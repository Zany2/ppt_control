# WPS / PowerPoint 控制 API

用于远程控制 WPS 演示或 Microsoft PowerPoint 的 HTTP API 服务。

通过 COM 自动化接口操控本地 WPS/PowerPoint 应用，支持打开文件、启动放映、翻页、跳转、自动翻页、退出放映等功能，适用于演讲、演出等场景下的自动化幻灯片控制。

## 环境要求

- **操作系统：** Windows
- **Python：** 3.10+
- **办公软件：** WPS Office 或 Microsoft Office（至少安装其一）

## 安装

```bash
# 克隆项目
git clone <仓库地址>
cd ppt_control

# 创建虚拟环境
python -m venv .venv
.venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt
```

## 启动服务

```bash
python ppt_control.py
```

启动后服务监听 `http://0.0.0.0:8000`，默认自动启动 WPS（若未安装则回退到 PowerPoint）。

- **Swagger 文档：** http://127.0.0.1:8000/docs
- **ReDoc 文档：** http://127.0.0.1:8000/redoc

## 响应格式

所有接口统一返回以下 JSON 格式：

```json
{
  "code": 20000,
  "message": "操作结果描述"
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| `code` | int | 状态码，`20000` 成功，`50000` 失败 |
| `message` | string | 成功时为操作结果描述，失败时为错误提示信息 |

## 接口列表

### 应用管理

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/ppt/start_app` | 启动或重启 WPS / PowerPoint 应用 |
| GET | `/api/ppt/app_info` | 获取当前运行的应用信息（名称、版本、状态） |
| GET | `/api/status` | 获取系统整体状态 |
| POST | `/api/ppt/exit_app` | 优雅关闭应用（通过 COM） |
| POST | `/api/ppt/force_close` | 强制关闭应用进程 |

### 文件操作

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/ppt/open` | 打开指定路径的 PPT 文件 |
| POST | `/api/ppt/close` | 关闭演示文稿（不关闭应用） |

### 放映控制

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/ppt/start` | 启动幻灯片放映 |
| GET | `/api/ppt/is_ready` | 检查是否已进入放映模式 |
| GET | `/api/ppt/current` | 获取当前幻灯片位置 |
| POST | `/api/ppt/next` | 下一张幻灯片 |
| POST | `/api/ppt/prev` | 上一张幻灯片 |
| POST | `/api/ppt/goto` | 跳转到指定幻灯片 |
| POST | `/api/ppt/blank` | 黑屏/白屏/恢复 |
| POST | `/api/ppt/auto_play` | 根据时间线自动翻页 |
| POST | `/api/ppt/stop_auto_play` | 停止自动翻页 |
| POST | `/api/ppt/exit_show` | 退出放映模式 |

### 媒体信息

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/media/info` | 获取当前幻灯片中的媒体信息（视频/音频） |

## 接口详细说明

### POST /api/ppt/start_app

启动或重启 WPS / PowerPoint 应用。

**请求参数：**

```json
{
  "prefer": "wps"
}
```

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `prefer` | string | `"wps"` | 首选应用：`wps` 或 `ppt` |

### POST /api/ppt/open

打开指定 PPT 文件。

**请求参数：**

```json
{
  "file_path": "C:\Users\xxx\Desktop\demo.pptx"
}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `file_path` | string | 是 | PPT 文件的完整路径 |

### POST /api/ppt/goto

跳转到指定幻灯片。

**请求参数：**

```json
{
  "slide": 3
}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `slide` | int | 是 | 目标幻灯片编号（从 1 开始） |

### POST /api/ppt/auto_play

根据时间线自动翻页，适合配合音乐或演讲节奏控制 PPT。

**请求参数：**

```json
{
  "timeline": [[38.0, 2, 1.5], [50.0, 3, 1.2]],
  "lead_time": 0.0,
  "auto_exit": false
}
```

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `timeline` | array | 必填 | 时间线数组 `[[时间点, 翻页次数, 翻页间隔秒], ...]` |
| `lead_time` | float | `0.0` | 提前触发时间（秒） |
| `auto_exit` | bool | `false` | 播放完成后是否自动退出放映 |

**timeline 示例说明：**

`[[38.0, 2, 1.5], [50.0, 3, 1.2]]` 表示：
- 第 38 秒：点击 2 次，每次间隔 1.5 秒
- 第 50 秒：点击 3 次，每次间隔 1.2 秒

### POST /api/ppt/stop_auto_play

停止正在运行的自动翻页任务。此接口不受串行化中间件限制，可在自动翻页运行期间随时调用。

**无需请求参数。**

### POST /api/ppt/blank

在放映模式下切换黑屏、白屏或恢复正常显示。

**请求参数：**

```json
{
  "action": "black"
}
```

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `action` | string | `"black"` | `black`（黑屏）、`white`（白屏）、`resume`（恢复正常） |

### POST /api/ppt/close

关闭当前打开的演示文稿，但不关闭 WPS / PowerPoint 应用。如果正在放映会先退出放映模式。

**无需请求参数。**

## 打包为 EXE

```bash
pyinstaller --onefile --name ppt_control ^
  --hidden-import=comtypes.stream ^
  --hidden-import=uvicorn.logging ^
  --hidden-import=uvicorn.loops ^
  --hidden-import=uvicorn.loops.auto ^
  --hidden-import=uvicorn.protocols ^
  --hidden-import=uvicorn.protocols.http ^
  --hidden-import=uvicorn.protocols.http.auto ^
  --hidden-import=uvicorn.protocols.websockets ^
  --hidden-import=uvicorn.protocols.websockets.auto ^
  --hidden-import=uvicorn.lifespan ^
  --hidden-import=uvicorn.lifespan.on ^
  ppt_control.py
```

打包完成后在 `dist\` 目录下生成 `ppt_control.exe`，可直接复制到任意位置双击运行，目标机器无需安装 Python 环境。

## 项目结构

```
ppt_control/
├── ppt_control.py       # 主程序（FastAPI 服务）
├── requirements.txt     # Python 依赖
└── README.md            # 项目说明
```

## 技术说明

- **COM 自动化：** 通过 `comtypes` 库操作 WPS (`Kwpp.Application`) 或 PowerPoint (`PowerPoint.Application`)
- **线程安全：** 串行化中间件 + RLock + 线程本地 COM 初始化，确保 COM 对象不被多线程并发访问
- **容错机制：** COM 对象有效性检测、"被呼叫方拒绝"错误自动重试、进程存活双重检测
- **应用优先级：** 默认优先使用 WPS，失败自动回退到 PowerPoint
