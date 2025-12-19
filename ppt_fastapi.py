#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time    : 2025/12/19
# @File    : ppt_fastapi.py
# @Description : WPS / PowerPoint 控制 API 服务（FastAPI 版本）

"""
WPS / PowerPoint 幻灯片控制 API 服务 (FastAPI)
====================================

接口说明：
    - POST /api/ppt/start_app       手动启动或重启 WPS / PowerPoint 应用
    - GET  /api/ppt/app_info        获取当前运行软件信息
    - GET  /api/status              获取当前系统整体状态
    - POST /api/ppt/open            打开指定 PPT 文件
    - POST /api/ppt/start           启动幻灯片放映
    - GET  /api/ppt/is_ready        检查是否已进入放映模式
    - POST /api/ppt/next            切换到下一张幻灯片
    - POST /api/ppt/prev            切换到上一张幻灯片
    - POST /api/ppt/goto            跳转到指定幻灯片
    - POST /api/ppt/auto_play       根据时间线自动翻页（参数: timeline, lead_time, auto_exit）
    - POST /api/ppt/exit_show       退出放映模式
    - POST /api/ppt/exit_app        优雅关闭 WPS / PowerPoint 应用（通过COM）
    - POST /api/ppt/force_close     强制关闭 WPS / PowerPoint 应用进程
    - GET  /api/media/info          获取当前幻灯片中的媒体信息（视频/音频）

依赖安装：
    pip install fastapi uvicorn comtypes pywin32 psutil pydantic requests
"""

import comtypes.client as com
import comtypes
import psutil
from fastapi import FastAPI, HTTPException, Body, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field
from starlette.middleware.base import BaseHTTPMiddleware
from typing import Optional, List, Union, Any
from pathlib import Path
import traceback
import atexit
import time
import threading


# ==========================================================
# 中间件定义
# ==========================================================
class SerializationMiddleware(BaseHTTPMiddleware):
    """
    串行化中间件 - 确保请求按顺序处理

    说明:
        由于COM对象具有线程亲和性（只能在创建它的线程中使用），
        此中间件确保所有请求串行处理，避免多线程访问COM对象导致的错误。
    """
    async def dispatch(self, request: Request, call_next):
        # 获取信号量，确保同一时间只有一个请求在处理
        _request_semaphore.acquire()
        try:
            response = await call_next(request)
            return response
        finally:
            _request_semaphore.release()


# ==========================================================
# Pydantic 模型定义
# ==========================================================
class StartAppRequest(BaseModel):
    """启动应用请求模型"""
    prefer: str = Field(default="wps", description="首选应用: wps 或 ppt")


class OpenPPTRequest(BaseModel):
    """打开PPT请求模型"""
    file_path: str = Field(..., description="PPT 文件路径")


class GotoSlideRequest(BaseModel):
    """跳转幻灯片请求模型"""
    slide: int = Field(..., description="目标幻灯片编号", ge=1)


class AutoPlayRequest(BaseModel):
    """自动播放请求模型"""
    timeline: List[List[Union[float, int]]] = Field(..., description="时间线数组")
    lead_time: float = Field(default=0.0, description="提前时间")
    auto_exit: bool = Field(default=False, description="是否自动退出")


class ResponseModel(BaseModel):
    """标准响应模型"""
    status: str = Field(..., description="响应状态: success 或 error")
    message: Optional[str] = Field(None, description="响应消息")
    data: Optional[Any] = Field(None, description="响应数据")


# ==========================================================
# 全局变量定义
# ==========================================================
app = FastAPI(
    title="WPS / PowerPoint 控制 API",
    description="用于远程控制 WPS 演示或 PowerPoint 的 API 服务",
    version="2.0.0"
)

# 添加串行化中间件，确保请求按顺序处理（避免COM对象多线程访问问题）
app.add_middleware(SerializationMiddleware)

# 配置：是否在启动时自动启动WPS/PowerPoint
AUTO_START_APP = True  # 设置为 False 可禁用自动启动

ppt_app = None               # WPS / PowerPoint 应用对象
presentation = None          # 当前打开的演示文稿对象
slide_show = None            # 当前放映窗口对象
current_ppt_path = None      # 当前打开的 PPT 文件路径
use_wps = False              # 是否使用 WPS（True=使用WPS, False=使用PowerPoint）

# 线程本地存储，用于跟踪每个线程的COM初始化状态
_thread_local = threading.local()

# 线程锁，用于保护全局COM对象的访问
_com_lock = threading.RLock()

# 全局信号量，确保同一时间只有一个请求在处理
_request_semaphore = threading.Semaphore(1)


# ==========================================================
# 常量定义
# ==========================================================
class PPTConstants:
    """
    幻灯片常量定义类

    常量说明：
        msoMedia: 媒体类型标识 (16表示媒体对象)
        ppMediaTypeSound: 音频媒体类型 (1表示声音)
        ppMediaTypeMovie: 视频媒体类型 (3表示视频/电影)
    """
    msoMedia = 16
    ppMediaTypeSound = 1
    ppMediaTypeMovie = 3


# ==========================================================
# 系统辅助函数
# ==========================================================
def is_process_running(name_keywords):
    """
    检测指定关键字的进程是否存在

    Args:
        name_keywords: 进程名关键字列表，如 ["wpp.exe", "wps"]

    Returns:
        bool: 如果找到匹配的进程返回 True，否则返回 False
    """
    for proc in psutil.process_iter(attrs=["name"]):
        name = proc.info["name"]
        if name and any(keyword.lower() in name.lower() for keyword in name_keywords):
            return True
    return False


def is_com_alive(obj):
    """
    检测 COM 对象是否仍然有效（深度检测）

    Args:
        obj: COM 对象实例

    Returns:
        bool: 如果 COM 对象有效返回 True，否则返回 False

    说明:
        通过访问 Visible 和 Presentations.Count 属性来验证对象有效性
    """
    if obj is None:
        return False
    try:
        _ = obj.Visible
        _ = obj.Presentations.Count
        return True
    except Exception:
        return False


def is_wps_alive():
    """
    检测 WPS 是否仍在运行（进程+COM双重检测）

    Returns:
        bool: WPS 进程存在且 COM 对象有效时返回 True
    """
    global ppt_app
    if not is_process_running(["wpp.exe", "wps"]):
        return False
    return is_com_alive(ppt_app)


def is_ppt_alive():
    """
    检测 PowerPoint 是否仍在运行（进程+COM双重检测）

    Returns:
        bool: PowerPoint 进程存在且 COM 对象有效时返回 True
    """
    global ppt_app
    if not is_process_running(["POWERPNT.EXE", "powerpoint"]):
        return False
    return is_com_alive(ppt_app)


# ==========================================================
# COM 初始化与应用检测
# ==========================================================
def init_com():
    """
    初始化 COM 环境（线程安全）

    Returns:
        bool: 初始化成功返回 True，失败返回 False

    说明:
        使用 comtypes.CoInitialize() 初始化 COM 环境，
        每个线程都需要独立初始化。使用线程本地存储确保
        每个线程只初始化一次。
    """
    try:
        # 检查当前线程是否已初始化
        if not getattr(_thread_local, 'initialized', False):
            comtypes.CoInitialize()
            _thread_local.initialized = True
            print(f"[INFO] COM 初始化成功 (线程ID: {threading.current_thread().ident})")
        return True
    except Exception as e:
        print(f"[ERROR] 初始化 COM 失败: {e}")
        return False


def ensure_app():
    """
    确保 WPS 或 PowerPoint 应用可用。
    若对象失效或进程关闭则重新启动。
    """
    global ppt_app, use_wps
    init_com()

    # 若已有对象且可用，直接返回
    if ppt_app and is_com_alive(ppt_app):
        return True

    # 若对象存在但无效，清理旧对象
    if ppt_app:
        try:
            ppt_app.Quit()
        except:
            pass
        ppt_app = None

    # 检测现有进程并尝试连接
    wps_running = is_process_running(["wpp.exe", "wps"])
    ppt_running = is_process_running(["POWERPNT.EXE", "powerpoint"])

    try:
        if wps_running:
            print("[INFO] 检测到 WPS 进程，尝试连接中...")
            ppt_app = com.GetActiveObject("Kwpp.Application")
            ppt_app.Visible = True
            use_wps = True
            return True
        elif ppt_running:
            print("[INFO] 检测到 PowerPoint 进程，尝试连接中...")
            ppt_app = com.GetActiveObject("PowerPoint.Application")
            ppt_app.Visible = True
            use_wps = False
            return True
    except Exception:
        pass

    # 若无任何实例则重新启动
    try:
        print("[INFO] 启动 WPS 演示 (Kwpp.Application)")
        ppt_app = com.CreateObject("Kwpp.Application")
        ppt_app.Visible = True
        use_wps = True
        print("[INFO] 已成功启动 WPS 演示")
        return True
    except Exception:
        print("[WARN] 启动 WPS 失败，尝试使用 PowerPoint")
        try:
            ppt_app = com.CreateObject("PowerPoint.Application")
            ppt_app.Visible = True
            use_wps = False
            print("[INFO] 已成功启动 PowerPoint")
            return True
        except Exception as e:
            print(f"[ERROR] 启动失败: {e}")
            return False


def ensure_presentation():
    """
    确保演示文稿对象有效（线程安全版本，带重试机制）

    Returns:
        bool: 演示文稿对象有效返回 True，否则返回 False

    说明:
        通过多重检查确保演示文稿对象的有效性：
        1. 检查全局 presentation 对象
        2. 检查 ppt_app.Presentations 集合
        3. 如果全局对象失效但应用中有演示文稿，则尝试恢复引用
        4. 对于"被呼叫方拒绝"错误，进行短暂等待后重试
    """
    global presentation, ppt_app

    with _com_lock:
        # 首先检查全局 presentation 对象
        try:
            if presentation:
                _ = presentation.Name
                return True
        except:
            presentation = None

        # 如果全局对象失效，尝试从应用中获取当前打开的演示文稿
        max_retries = 3
        for attempt in range(max_retries):
            try:
                if ppt_app and ppt_app.Presentations.Count > 0:
                    # 获取最近打开的演示文稿（通常是第一个）
                    presentation = ppt_app.Presentations.Item(1)
                    _ = presentation.Name  # 验证对象有效性
                    print(f"[INFO] 已从应用中恢复演示文稿引用: {presentation.Name}")
                    return True
                else:
                    break  # 没有打开的演示文稿，无需重试
            except Exception as e:
                error_code = getattr(e, 'args', [None])[0] if hasattr(e, 'args') else None

                # -2147418111 是 RPC_E_CALL_REJECTED（被呼叫方拒绝接收呼叫）
                if error_code == -2147418111 and attempt < max_retries - 1:
                    print(f"[WARN] COM对象忙碌，等待重试 ({attempt + 1}/{max_retries})...")
                    time.sleep(0.8)  # 等待800ms后重试，给WPS更多时间
                    continue
                else:
                    if attempt == 0:  # 只在第一次失败时打印详细错误
                        print(f"[WARN] 无法从应用中获取演示文稿: {e}")
                    presentation = None
                    break

        return False


def ensure_slideshow():
    """
    确保放映窗口对象有效

    Returns:
        bool: 放映窗口对象有效返回 True，否则返回 False

    说明:
        通过访问 View 属性检查放映窗口是否有效，
        如果无效则将全局变量 slide_show 设置为 None
    """
    global slide_show
    try:
        if slide_show:
            _ = slide_show.View
            return True
    except:
        slide_show = None
    return False


@atexit.register
def cleanup():
    """
    程序退出时释放 COM 环境

    说明:
        在主线程退出时尝试释放COM环境。
        由于使用了线程本地存储，每个线程的COM环境会独立管理。
    """
    try:
        if getattr(_thread_local, 'initialized', False):
            comtypes.CoUninitialize()
            print("[INFO] COM 环境已释放")
    except:
        pass


# ==========================================================
# 辅助函数
# ==========================================================
def auto_start_app_background():
    """
    后台任务：通过HTTP请求自动启动WPS/PowerPoint应用

    说明:
        等待服务启动完成后，通过发送HTTP请求到 /api/ppt/start_app 接口
        来触发应用启动。这样可以确保COM对象在工作线程中创建。
    """
    try:
        import requests

        print("[INFO] 后台任务：等待服务启动完成...")
        time.sleep(2)  # 等待服务完全启动

        print("[INFO] 后台任务：正在自动启动 WPS/PowerPoint...")
        response = requests.post(
            "http://127.0.0.1:8000/api/ppt/start_app",
            json={"prefer": "wps"},
            timeout=10
        )

        if response.status_code == 200:
            result = response.json()
            print(f"[INFO] 自动启动成功: {result.get('message', '')}")
        else:
            print(f"[WARN] 自动启动请求返回状态码: {response.status_code}")
    except ImportError:
        print("[WARN] 需要安装 requests 库才能自动启动: pip install requests")
        print("[INFO] WPS/PowerPoint 将在第一次API调用时启动")
    except Exception as e:
        print(f"[WARN] 后台自动启动失败: {e}")
        print("[INFO] WPS/PowerPoint 将在第一次API调用时启动")


@app.on_event("startup")
async def startup_event():
    """
    服务启动事件

    说明:
        如果启用了AUTO_START_APP，则在后台发送HTTP请求自动启动WPS/PowerPoint
    """
    if AUTO_START_APP:
        print("[INFO] 启动事件：已安排自动启动任务")
        # 在独立线程中发送HTTP请求，确保COM对象在工作线程中创建
        thread = threading.Thread(target=auto_start_app_background, daemon=True)
        thread.start()
    else:
        print("[INFO] 自动启动已禁用，WPS/PowerPoint 将在第一次API调用时启动")


def get_media_shapes(slide):
    """
    获取当前幻灯片中的媒体对象信息

    Args:
        slide: 幻灯片对象

    Returns:
        list: 媒体对象列表，每个元素包含 name 和 type 字段
              type 可能的值: "video" 或 "audio"

    说明:
        遍历幻灯片中的所有形状，识别媒体类型（视频/音频）
    """
    media_list = []
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes.Item(i)
        if shape.Type == PPTConstants.msoMedia:
            media_type = getattr(shape, "MediaType", None)
            if media_type == PPTConstants.ppMediaTypeMovie:
                media_list.append({"name": shape.Name, "type": "video"})
            elif media_type == PPTConstants.ppMediaTypeSound:
                media_list.append({"name": shape.Name, "type": "audio"})
    return media_list


# ==========================================================
# FastAPI 路由定义
# ==========================================================
@app.post("/api/ppt/start_app", response_model=ResponseModel)
def start_app(request: StartAppRequest = Body(default=StartAppRequest())):
    """
    手动启动或重启 WPS / PowerPoint 应用。

    参数:
        prefer: "wps" 或 "ppt"（默认 "wps"）
    """
    global ppt_app, use_wps
    prefer = request.prefer.lower()

    # 确保COM环境已初始化
    init_com()

    try:
        wps_alive = is_wps_alive()
        ppt_alive = is_ppt_alive()

        # 如果当前对象失效则清理
        if ppt_app:
            if (use_wps and not wps_alive) or (not use_wps and not ppt_alive):
                print("[WARN] 检测到旧对象失效，清理中...")
                try:
                    ppt_app.Quit()
                except:
                    pass
                ppt_app = None

        # 若对象仍可用则直接返回
        if ppt_app and is_com_alive(ppt_app):
            app_name = "WPS" if use_wps else "PowerPoint"
            return ResponseModel(
                status="success",
                message=f"{app_name} 已在运行中",
                data={"use_wps": use_wps}
            )

        # 启动新实例
        if prefer == "ppt":
            ppt_app = com.CreateObject("PowerPoint.Application")
            ppt_app.Visible = True
            use_wps = False
            app_name = "PowerPoint"
        else:
            try:
                ppt_app = com.CreateObject("Kwpp.Application")
                ppt_app.Visible = True
                use_wps = True
                app_name = "WPS 演示"
            except Exception:
                ppt_app = com.CreateObject("PowerPoint.Application")
                ppt_app.Visible = True
                use_wps = False
                app_name = "PowerPoint"

        return ResponseModel(
            status="success",
            message=f"{app_name} 已启动",
            data={"use_wps": use_wps}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/ppt/app_info", response_model=ResponseModel)
def app_info():
    """
    获取当前运行的应用信息

    功能说明:
        返回当前 WPS 或 PowerPoint 应用的详细信息

    返回:
        ResponseModel: 包含应用信息的响应对象
            - running: 应用是否正在运行
            - app: 应用名称 ("WPS" 或 "PowerPoint")
            - version: 应用版本号
            - use_wps: 是否使用 WPS
    """
    global ppt_app, use_wps

    # 确保COM环境已初始化
    init_com()

    try:
        if not ppt_app:
            return ResponseModel(
                status="success",
                data={"running": False, "app": None, "version": None}
            )

        app_name = "WPS" if use_wps else "PowerPoint"
        process_alive = is_wps_alive() if use_wps else is_ppt_alive()

        # 安全地获取版本信息
        version = None
        try:
            version = ppt_app.Version
        except Exception as e:
            print(f"[WARN] 获取版本信息失败: {e}")

        return ResponseModel(
            status="success",
            data={
                "running": process_alive and is_com_alive(ppt_app),
                "app": app_name,
                "version": version,
                "use_wps": use_wps
            }
        )
    except Exception as e:
        # 如果是COM未初始化错误，返回友好的错误信息
        error_msg = str(e)
        if "CoInitialize" in error_msg:
            raise HTTPException(
                status_code=503,
                detail="COM环境初始化失败，请稍后重试或调用 /api/ppt/start_app 接口"
            )
        raise HTTPException(status_code=500, detail=error_msg)


@app.get("/api/status", response_model=ResponseModel)
def status():
    """
    获取系统运行状态

    功能说明:
        返回系统整体运行状态，包括应用、演示文稿、放映等信息

    返回:
        ResponseModel: 包含系统状态的响应对象
            - use_wps: 是否使用 WPS
            - app_ready: 应用对象是否就绪
            - wps_process_alive: WPS 进程是否存活
            - ppt_process_alive: PowerPoint 进程是否存活
            - presentation_open: 演示文稿是否已打开
            - slideshow_running: 放映是否正在运行
            - current_ppt: 当前打开的 PPT 文件路径
    """
    # 确保COM环境已初始化
    init_com()

    # 安全地检查演示文稿和放映状态
    presentation_open = False
    slideshow_running = False

    try:
        presentation_open = ensure_presentation()
    except Exception as e:
        print(f"[WARN] 检查演示文稿状态失败: {e}")

    try:
        slideshow_running = ensure_slideshow()
    except Exception as e:
        print(f"[WARN] 检查放映状态失败: {e}")

    return ResponseModel(
        status="success",
        data={
            "use_wps": use_wps,
            "app_ready": ppt_app is not None,
            "wps_process_alive": is_process_running(["wpp.exe", "wps"]),
            "ppt_process_alive": is_process_running(["POWERPNT.EXE", "powerpoint"]),
            "presentation_open": presentation_open,
            "slideshow_running": slideshow_running,
            "current_ppt": current_ppt_path
        }
    )


@app.post("/api/ppt/open", response_model=ResponseModel)
def open_ppt(request: OpenPPTRequest):
    """
    打开指定 PPT 文件

    参数:
        request.file_path: PPT 文件的完整路径

    功能说明:
        1. 验证文件路径是否有效
        2. 确保应用已启动
        3. 关闭当前打开的演示文稿（如果有）
        4. 打开新的 PPT 文件

    返回:
        ResponseModel: 包含打开结果的响应对象
            - slides: 幻灯片总数
            - use_wps: 使用的应用类型

    异常:
        HTTPException: 当文件路径无效或打开失败时抛出错误
    """
    global presentation, current_ppt_path
    ppt_path = request.file_path

    if not ppt_path or not Path(ppt_path).exists():
        raise HTTPException(status_code=400, detail="文件路径无效")

    # 确保COM环境已初始化
    init_com()

    if not ensure_app():
        raise HTTPException(status_code=500, detail="无法启动 WPS / PowerPoint")

    try:
        with _com_lock:
            if ensure_presentation():
                try:
                    presentation.Close()
                except:
                    pass

            presentation = ppt_app.Presentations.Open(ppt_path, WithWindow=True)
            current_ppt_path = ppt_path

            # 验证演示文稿已成功打开
            slides_count = presentation.Slides.Count
            print(f"[INFO] 成功打开演示文稿: {Path(ppt_path).name}, 共 {slides_count} 张幻灯片")

            # 等待WPS/PowerPoint完全加载文件（避免后续操作遇到COM对象忙碌）
            # 根据幻灯片数量动态调整等待时间：基础1.5秒 + 每1张幻灯片额外30ms
            wait_time = 1.5 + slides_count * 0.03
            wait_time = min(wait_time, 6.0)  # 最多等待6秒
            time.sleep(wait_time)
            print(f"[INFO] 演示文稿加载完成，已就绪（等待 {wait_time:.2f}s）")

            return ResponseModel(
                status="success",
                message=f"成功打开 {Path(ppt_path).name}",
                data={"slides": slides_count, "use_wps": use_wps}
            )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/ppt/start", response_model=ResponseModel)
def start_show():
    """
    启动幻灯片放映

    前置条件:
        必须先打开 PPT 文件

    功能说明:
        启动幻灯片放映模式，从第一张幻灯片开始播放

    返回:
        ResponseModel: 包含放映信息的响应对象
            - current_slide: 当前幻灯片编号

    异常:
        HTTPException: 当未打开 PPT 或启动失败时抛出错误
    """
    global slide_show

    # 确保COM环境已初始化
    init_com()

    with _com_lock:
        # 增强的演示文稿检测
        if not ensure_presentation():
            raise HTTPException(status_code=400, detail="未打开任何 PPT，请先调用 /api/ppt/open 接口打开文件")

        try:
            # 验证演示文稿对象仍然有效
            ppt_name = presentation.Name
            slides_count = presentation.Slides.Count
            print(f"[INFO] 准备启动放映: {ppt_name}, 共 {slides_count} 张幻灯片")

            # 短暂等待确保演示文稿完全就绪
            time.sleep(0.5)

            # 启动放映（带重试机制）
            max_retries = 3
            last_error = None
            current_slide = None

            for attempt in range(max_retries):
                try:
                    slide_show = presentation.SlideShowSettings.Run()
                    current_slide = slide_show.View.CurrentShowPosition
                    print(f"[INFO] 放映已启动，当前在第 {current_slide} 张幻灯片")
                    break  # 成功则跳出循环
                except Exception as e:
                    last_error = e
                    error_code = getattr(e, 'args', [None])[0] if hasattr(e, 'args') else None

                    # 如果是COM对象忙碌，等待后重试
                    if error_code == -2147418111 and attempt < max_retries - 1:
                        print(f"[WARN] 启动放映时COM对象忙碌，等待重试 ({attempt + 1}/{max_retries})...")
                        time.sleep(1.0)  # 等待1秒
                        continue
                    elif attempt < max_retries - 1:
                        # 其他错误也重试
                        print(f"[WARN] 启动放映失败，重试 ({attempt + 1}/{max_retries}): {e}")
                        time.sleep(0.8)
                        continue
                    else:
                        # 最后一次重试失败，抛出异常
                        raise last_error

            # 验证启动成功
            if current_slide is None:
                raise HTTPException(status_code=500, detail="启动放映失败，请重试")

            return ResponseModel(
                status="success",
                message="幻灯片放映已启动",
                data={
                    "current_slide": current_slide,
                    "total_slides": slides_count,
                    "ppt_name": ppt_name
                }
            )
        except Exception as e:
            traceback.print_exc()
            raise HTTPException(status_code=500, detail=f"启动放映失败: {str(e)}")


@app.get("/api/ppt/is_ready", response_model=ResponseModel)
def is_ready():
    """检查是否已进入放映模式"""
    # 确保COM环境已初始化
    init_com()

    try:
        with _com_lock:
            if not ensure_app():
                raise HTTPException(status_code=400, detail="应用未运行")
            if not ensure_presentation():
                raise HTTPException(status_code=400, detail="PPT 未打开")
            if not ensure_slideshow():
                raise HTTPException(status_code=400, detail="未进入放映模式")

            view = slide_show.View
            slide_num = view.CurrentShowPosition
            total_slides = presentation.Slides.Count

            return ResponseModel(
                status="success",
                data={
                    "ready": True,
                    "app": "WPS" if use_wps else "PowerPoint",
                    "current_slide": slide_num,
                    "total_slides": total_slides,
                    "message": "放映模式已启动"
                }
            )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/ppt/next", response_model=ResponseModel)
def next_slide():
    """
    切换到下一张幻灯片

    前置条件:
        必须处于放映模式

    功能说明:
        在放映模式下切换到下一张幻灯片

    返回:
        ResponseModel: 包含当前幻灯片编号的响应对象
    """
    if not ensure_slideshow():
        raise HTTPException(status_code=400, detail="放映未启动")
    try:
        slide_show.View.Next()
        pos = slide_show.View.CurrentShowPosition
        return ResponseModel(status="success", message=f"切换到第 {pos} 张")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/ppt/prev", response_model=ResponseModel)
def prev_slide():
    """
    切换到上一张幻灯片

    前置条件:
        必须处于放映模式

    功能说明:
        在放映模式下切换到上一张幻灯片

    返回:
        ResponseModel: 包含当前幻灯片编号的响应对象
    """
    if not ensure_slideshow():
        raise HTTPException(status_code=400, detail="放映未启动")
    try:
        slide_show.View.Previous()
        pos = slide_show.View.CurrentShowPosition
        return ResponseModel(status="success", message=f"返回到第 {pos} 张")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/ppt/goto", response_model=ResponseModel)
def goto_slide(request: GotoSlideRequest):
    """
    跳转到指定幻灯片

    参数:
        request.slide: 目标幻灯片编号（从 1 开始）

    前置条件:
        必须处于放映模式

    功能说明:
        在放映模式下直接跳转到指定编号的幻灯片

    返回:
        ResponseModel: 包含跳转结果的响应对象

    异常:
        HTTPException: 当幻灯片编号超出范围时抛出 400 错误
    """
    if not ensure_slideshow():
        raise HTTPException(status_code=400, detail="放映未启动")
    try:
        slide_num = request.slide
        total = presentation.Slides.Count
        if slide_num < 1 or slide_num > total:
            raise HTTPException(
                status_code=400,
                detail=f"幻灯片编号超出范围 (1-{total})"
            )
        slide_show.View.GotoSlide(slide_num)
        return ResponseModel(
            status="success",
            message=f"已跳转到第 {slide_num} 张幻灯片"
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/ppt/exit_show", response_model=ResponseModel)
def exit_show():
    """
    退出放映模式

    功能说明:
        退出当前的幻灯片放映模式，返回编辑视图
        不会关闭演示文稿或应用程序

    返回:
        ResponseModel: 包含执行结果的响应对象

    说明:
        如果当前未在放映模式，也会返回成功
    """
    global slide_show
    if not ensure_slideshow():
        return ResponseModel(status="success", message="当前未在放映模式")
    try:
        slide_show.View.Exit()
        slide_show = None
        return ResponseModel(status="success", message="已退出放映模式")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/ppt/exit_app", response_model=ResponseModel)
def exit_app():
    """
    优雅关闭 WPS / PowerPoint 应用（通过 COM）

    功能说明:
        1. 退出放映模式（如果正在放映）
        2. 关闭当前演示文稿（如果已打开）
        3. 退出应用程序
        4. 清理所有全局对象引用

    返回:
        ResponseModel: 包含执行结果的响应对象

    异常:
        HTTPException: 当操作失败时抛出 500 错误
    """
    global ppt_app, presentation, slide_show
    try:
        if slide_show:
            slide_show.View.Exit()
            slide_show = None
        if presentation:
            presentation.Close()
            presentation = None
        if ppt_app:
            ppt_app.Quit()
            ppt_app = None
        return ResponseModel(status="success", message="WPS / PowerPoint 已优雅退出")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/ppt/force_close", response_model=ResponseModel)
def force_close_app():
    """
    强制关闭 WPS / PowerPoint 应用进程

    功能说明:
        直接终止 WPS/PowerPoint 相关进程，不进行优雅退出
        适用于应用无响应或 COM 调用失败的情况

    警告:
        此方法会强制结束进程，可能导致未保存的数据丢失
        建议优先使用 /api/ppt/exit_app 进行优雅关闭

    返回:
        ResponseModel: 包含执行结果和被关闭的进程列表

    异常:
        HTTPException: 当操作失败时抛出 500 错误
    """
    global ppt_app, presentation, slide_show

    try:
        # 清理全局对象引用
        slide_show = None
        presentation = None
        ppt_app = None

        killed_processes = []
        wps_keywords = ["wpp.exe", "wps"]
        ppt_keywords = ["POWERPNT.EXE", "powerpoint"]

        # 一次遍历，同时查找 WPS 和 PowerPoint 进程
        for proc in psutil.process_iter(attrs=["pid", "name"]):
            proc_name = proc.info["name"]
            if not proc_name:
                continue

            # 判断是否为 WPS 或 PowerPoint 进程
            is_wps = any(keyword.lower() in proc_name.lower() for keyword in wps_keywords)
            is_ppt = any(keyword.lower() in proc_name.lower() for keyword in ppt_keywords)

            if is_wps or is_ppt:
                try:
                    proc_type = "WPS" if is_wps else "PowerPoint"
                    proc.kill()
                    killed_processes.append({
                        "pid": proc.info["pid"],
                        "name": proc_name,
                        "type": proc_type
                    })
                    print(f"[INFO] 已强制关闭 {proc_type} 进程: {proc_name} (PID: {proc.info['pid']})")
                except Exception as e:
                    print(f"[WARN] 关闭进程 {proc_name} 失败: {e}")

        if killed_processes:
            # 统计各类型进程数量
            wps_count = sum(1 for p in killed_processes if p.get("type") == "WPS")
            ppt_count = sum(1 for p in killed_processes if p.get("type") == "PowerPoint")

            summary = []
            if wps_count > 0:
                summary.append(f"WPS {wps_count} 个")
            if ppt_count > 0:
                summary.append(f"PowerPoint {ppt_count} 个")

            message = f"已强制关闭 {', '.join(summary)} 进程"

            return ResponseModel(
                status="success",
                message=message,
                data={"killed_processes": killed_processes, "total": len(killed_processes)}
            )
        else:
            return ResponseModel(
                status="success",
                message="未找到运行中的 WPS / PowerPoint 进程"
            )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/media/info", response_model=ResponseModel)
def media_info():
    """
    获取当前幻灯片中的媒体信息

    功能说明:
        返回当前幻灯片中的所有媒体对象（视频/音频）信息

    前置条件:
        必须处于放映模式

    返回:
        ResponseModel: 包含媒体信息的响应对象
            - slide_number: 当前幻灯片编号
            - media_count: 媒体对象数量
            - media: 媒体对象列表 [{name, type}, ...]

    异常:
        HTTPException: 当未进入放映模式时抛出 400 错误
    """
    if not ensure_slideshow():
        raise HTTPException(status_code=400, detail="放映未启动")
    try:
        slide = slide_show.View.Slide
        media = get_media_shapes(slide)
        return ResponseModel(
            status="success",
            data={
                "slide_number": slide_show.View.CurrentShowPosition,
                "media_count": len(media),
                "media": media
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/ppt/auto_play", response_model=ResponseModel)
def auto_play(request: AutoPlayRequest):
    """
    自动翻页（节奏驱动版，无动画检测、无时间累积误差）

    参数:
        request.timeline: 时间线数组，格式如下:
            [[时间点, 翻页次数, 翻页间隔秒], ...]
            示例: [[38.0, 2, 1.5], [50.0, 3, 1.2]]
        request.lead_time: 提前时间（秒），默认 0.0
        request.auto_exit: 是否自动退出放映模式，默认 False

    前置条件:
        必须处于放映模式

    功能说明:
        根据时间线自动翻页，每个时间点独立执行
        每组时间点内点击多次，每次之间固定间隔
        不计算累计时间差，保证后续节点不会被前一组的耗时影响

    返回:
        ResponseModel: 包含执行结果和统计信息的响应对象

    异常:
        HTTPException: 当未进入放映模式或参数错误时抛出错误
    """
    try:
        # 1. 确认放映模式
        if not ensure_slideshow():
            raise HTTPException(status_code=400, detail="请先进入放映模式")

        # 2. 解析参数
        timeline = request.timeline
        lead_time = request.lead_time
        auto_exit = request.auto_exit

        if not timeline or not isinstance(timeline, list):
            raise HTTPException(
                status_code=400,
                detail="timeline 参数缺失或格式错误，应为二维数组"
            )

        print(f"[AUTO] 自动翻页启动，共 {len(timeline)} 个时间点")
        base_time = time.time()  # 程序启动时间（仅用于初始同步）

        # 3. 遍历每个时间节点
        for i, item in enumerate(timeline, start=1):
            if not isinstance(item, (list, tuple)) or len(item) < 1:
                print(f"[AUTO] 第 {i} 项格式错误，跳过: {item}")
                continue

            target_time = float(item[0])         # 节点触发时间
            flip_count = int(item[1]) if len(item) > 1 else 1
            flip_interval = float(item[2]) if len(item) > 2 else 1.0

            print(f"[AUTO] 等待 {target_time:.2f}s 到达执行点 (点击 {flip_count} 次, 间隔 {flip_interval}s)")

            # 独立等待到达节点时间（相对程序启动）
            while time.time() - base_time < (target_time - lead_time):
                time.sleep(0.05)

            print(f"[AUTO] 到达 {target_time:.2f}s，开始第 {i} 组点击动作")

            # 执行多次点击
            for j in range(flip_count):
                try:
                    if not slide_show:
                        raise HTTPException(status_code=400, detail="放映窗口已关闭")
                    print(time.time())
                    slide_show.View.Next()
                    print(f"    [AUTO] 点击 {j+1}/{flip_count}")

                    # 除最后一次外，每次都固定等待 flip_interval 秒
                    if j < flip_count - 1:
                        time.sleep(flip_interval)

                except HTTPException:
                    raise
                except Exception as e:
                    print(f"    [ERROR] 点击失败: {e}")
                    # 出错也等一段时间再继续
                    time.sleep(flip_interval)

        # 4. 自动退出放映
        if auto_exit and slide_show:
            print("[AUTO] 所有动作完成，退出放映模式")
            try:
                slide_show.View.Exit()
            except Exception as e:
                print(f"[WARN] 自动退出失败: {e}")

        print("[AUTO] 自动翻页任务完成")

        # 5. 返回执行结果
        return ResponseModel(
            status="success",
            message="自动翻页任务执行完毕（节奏驱动版）",
            data={
                "timeline": timeline,
                "lead_time": lead_time,
                "auto_exit": auto_exit,
                "total_points": len(timeline)
            }
        )

    except HTTPException:
        raise
    except Exception as e:
        print(f"[AUTO] 异常: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# ==========================================================
# 主入口
# ==========================================================
if __name__ == "__main__":
    import uvicorn

    print("=== WPS / PowerPoint 控制服务器 (FastAPI) ===")
    if AUTO_START_APP:
        print("[提示] 服务启动后将自动启动 WPS/PowerPoint（可通过修改 AUTO_START_APP 禁用）")
    else:
        print("[提示] 自动启动已禁用，WPS/PowerPoint 将在第一次API调用时启动")

    print("\n接口列表：")
    print("=" * 100)
    print("【应用管理】")
    print("  - POST /api/ppt/start_app      # 手动启动或重启 WPS 或 PowerPoint 应用（参数: prefer='wps'/'ppt'）")
    print("  - GET  /api/ppt/app_info       # 获取当前运行软件信息（版本、进程、COM 状态等）")
    print("  - GET  /api/status             # 获取当前系统整体状态（应用、演示文稿、放映等）")
    print("  - POST /api/ppt/exit_app       # 优雅关闭 WPS / PowerPoint 应用（通过 COM）")
    print("  - POST /api/ppt/force_close    # 强制关闭 WPS / PowerPoint 应用进程（适用于无响应情况）")
    print()
    print("【文件操作】")
    print("  - POST /api/ppt/open           # 打开指定 PPT 文件（参数: file_path）")
    print()
    print("【放映控制】")
    print("  - POST /api/ppt/start          # 启动幻灯片放映")
    print("  - GET  /api/ppt/is_ready       # 检查是否已进入放映模式")
    print("  - POST /api/ppt/next           # 切换到下一张幻灯片")
    print("  - POST /api/ppt/prev           # 切换到上一张幻灯片")
    print("  - POST /api/ppt/goto           # 跳转到指定幻灯片（参数: slide）")
    print("  - POST /api/ppt/auto_play      # 根据时间线自动翻页（参数: timeline, lead_time, auto_exit）")
    print("  - POST /api/ppt/exit_show      # 退出放映模式")
    print()
    print("【媒体信息】")
    print("  - GET  /api/media/info         # 获取当前幻灯片中的媒体信息（视频/音频）")
    print("=" * 100)
    print("\n服务器已启动：http://127.0.0.1:8000")
    print("API 文档：http://127.0.0.1:8000/docs")
    print("ReDoc 文档：http://127.0.0.1:8000/redoc")

    # 启动 Uvicorn 服务
    # 重要：
    # 1. 使用串行化中间件确保请求按顺序处理
    # 2. 使用workers=1确保单进程运行
    # 3. 重试机制处理COM对象忙碌的情况
    print("\n注意：已启用串行化中间件和重试机制，确保COM对象访问的线程安全")
    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000,
        workers=1  # 单工作进程
    )
