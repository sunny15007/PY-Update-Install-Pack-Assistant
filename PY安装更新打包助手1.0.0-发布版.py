import os
import sys
import time
import shutil
import subprocess
import webbrowser
import threading
import base64
import tempfile
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
from datetime import datetime


VERSION = "v1.0.0"
AUTHOR = "sunny15007"
APP_NAME = "PY更新安装打包助手"

REPORT_SYMBOL_COLORS = {
    "✅": "#1FA32B",
    "❌": "#E02020",
    "⚠️": "#E8A020",
    "💡": "#E0B000",
    "⏳": "#8A8A8A",
    "🔍": "#2E86DE",
    "📦": "#C8792B",
    "🔄": "#16A085",
    "🗑️": "#A04040",
    "⏱️": "#6C7A89",
    "📄": "#5A7FB0",
    "🖼️": "#7B68EE",
    "📁": "#D4A017",
    "📋": "#2E86C1",
    "🚀": "#E8590C",
    "⛔": "#C62828",
    "📥": "#3A7BD5",
}

if sys.platform == "win32":
    CREATE_NO_WINDOW = 0x08000000
else:
    CREATE_NO_WINDOW = 0

_IS_FROZEN = getattr(sys, "frozen", False) or ("__compiled__" in globals())
if _IS_FROZEN:
    PYTHON_CMD = ["python"]
else:
    _py_exe = sys.executable
    if os.path.basename(_py_exe).lower() == "pythonw.exe":
        _alt = os.path.join(os.path.dirname(_py_exe), "python.exe")
        if os.path.isfile(_alt):
            _py_exe = _alt
    PYTHON_CMD = [_py_exe]



SAFE_TEMP_DIR_NAMES = ("打包临时文件", "Nuitka临时文件")


def _get_protected_paths():
    home = os.path.abspath(os.path.expanduser("~"))
    desktop = os.path.abspath(os.path.join(home, "Desktop"))
    root = os.path.abspath(os.sep)
    protected = {root, home, desktop}
    if sys.platform == "win32":
        for letter in "CDEFGH":
            protected.add(os.path.abspath(f"{letter}:\\"))
    return protected


def is_safe_to_delete(path):
    if not path:
        return False, "路径为空"

    abs_path = os.path.abspath(path)
    norm_path = os.path.normpath(abs_path)

    if not os.path.isdir(norm_path):
        return False, "路径不存在或不是目录"

    dir_name = os.path.basename(norm_path)
    if dir_name not in SAFE_TEMP_DIR_NAMES:
        return False, f"目录名 '{dir_name}' 不在白名单 {SAFE_TEMP_DIR_NAMES} 中"

    parts = [p for p in norm_path.split(os.sep) if p]
    if len(parts) < 3:
        return False, f"路径层级过浅（{len(parts)} 层）：{norm_path}"

    protected = _get_protected_paths()
    for p in protected:
        try:
            p_norm = os.path.normpath(p)
            if norm_path == p_norm:
                return False, f"是受保护目录：{p_norm}"
            try:
                common = os.path.commonpath([norm_path, p_norm])
                if os.path.normpath(common) == norm_path and norm_path != p_norm:
                    return False, f"是受保护目录的父级：{p_norm}"
            except ValueError:
                pass
        except Exception:
            continue

    return True, "通过安全检查"


def safe_delete_dir(path):
    ok, reason = is_safe_to_delete(path)
    if not ok:
        return False, f"安全检查未通过，已拒绝删除。\n原因：{reason}\n路径：{path}"

    try:
        from send2trash import send2trash
    except ImportError:
        return False, (
            "未安装 send2trash，为安全起见本工具不会自动删除任何东西。\n\n"
            f"请手动删除：\n{path}\n\n"
            "（安装 send2trash 后，本工具可自动将临时文件移入回收站：\n"
            "  pip install send2trash ）"
        )

    try:
        send2trash(os.path.abspath(path))
        return True, f"已移入回收站：{path}"
    except Exception as e:
        return False, f"移入回收站失败：{e}\n路径：{path}"


def is_safe_nuitka_cache(path):
    if not path:
        return False, "路径为空"

    localappdata = os.environ.get("LOCALAPPDATA", "").strip()
    if not localappdata:
        return False, "未能获取系统 LOCALAPPDATA 路径"

    expect = os.path.normpath(os.path.abspath(os.path.join(localappdata, "Nuitka")))
    norm_path = os.path.normpath(os.path.abspath(path))

    if norm_path != expect:
        return False, f"路径不是 Nuitka 缓存目录（期望：{expect}）"

    if os.path.basename(norm_path).lower() != "nuitka":
        return False, f"目录名不是 Nuitka：{norm_path}"

    if not os.path.isdir(norm_path):
        return False, "路径不存在或不是目录"

    parts = [p for p in norm_path.split(os.sep) if p]
    if len(parts) < 3:
        return False, f"路径层级过浅（{len(parts)} 层）：{norm_path}"

    for p in _get_protected_paths():
        try:
            p_norm = os.path.normpath(p)
            if norm_path == p_norm:
                return False, f"是受保护目录：{p_norm}"
            try:
                common = os.path.commonpath([norm_path, p_norm])
                if os.path.normpath(common) == norm_path and norm_path != p_norm:
                    return False, f"是受保护目录的父级：{p_norm}"
            except ValueError:
                pass
        except Exception:
            continue

    return True, "通过安全检查"


def _get_sub_env():
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    env.pop("TCL_LIBRARY", None)
    env.pop("TK_LIBRARY", None)
    return env


LIB_DESC = {
    "openpyxl": "读写 Excel 表格(.xlsx)",
    "requests": "发送 HTTP 请求（爬虫/调用接口）",
    "beautifulsoup4": "解析网页 HTML（爬虫常用）",
    "PyQt5": "桌面软件界面开发",
    "PySide6": "桌面软件界面开发（Qt 官方）",
    "pyinstaller": "把 .py 打包成 exe 可执行文件",
    "Nuitka": "把 .py 编译成 exe（运行更快）",
    "py7zr": "压缩/解压 7z 文件",
    "psutil": "查看 CPU/内存/进程信息",
    "colorama": "让命令行文字显示颜色",
    "polib": "读写 .po 翻译文件（软件多语言）",
    "texttable": "在命令行里画表格",
    "pandas": "表格数据分析（Excel/CSV）",
    "numpy": "数值计算/矩阵运算（科学计算基础）",
    "matplotlib": "画图表（折线/柱状/饼图）",
    "lxml": "快速解析 HTML/XML",
    "selenium": "自动操作浏览器（爬虫/自动化测试）",
    "aiohttp": "异步 HTTP 请求（高并发爬虫）",
    "pyperclip": "读写系统剪贴板",
    "pillow": "图片处理（缩放/裁剪/转格式）",
    "opencv-python": "图像识别/视频处理（OpenCV）",
    "tqdm": "给循环加进度条",
    "schedule": "定时执行任务（每天/每小时）",
    "pyyaml": "读写 YAML 配置文件",
    "python-dotenv": "读取 .env 配置文件",
    "click": "快速写命令行工具",
    "Flask": "轻量级网站/接口开发",
    "Django": "完整网站开发框架",
    "jupyter": "网页版交互编程笔记本",
    "pytest": "编写/运行自动化测试",
    "black": "自动格式化 Python 代码",
    "xlrd": "读取旧版 Excel(.xls)",
    "xlwt": "写入旧版 Excel(.xls)",
    "pywin32": "调用 Windows 系统接口",
    "bottle": "轻量级微型 Web 框架",
    "cffi": "[依赖库] C 语言外部函数接口",
    "pythonnet": "Python 调用 .NET/C# 库",
    "clr_loader": "[依赖库] pythonnet .NET 加载器",
    "pycparser": "[依赖库] C 语言解析器",
    "pywebview": "用网页技术做桌面界面",
    "tkinterdnd2": "给 tkinter 添加拖拽功能",
    "win10toast": "Windows 桌面通知弹窗",
    "pypiwin32": "[依赖库] pywin32 旧版包名",
    "proxy_tools": "[依赖库] 代理工具",
    "taskbargap": "调整 Windows 任务栏图标间距",
    "pip": "包管理工具",
    "altgraph": "[依赖库] pyinstaller 依赖",
    "brotli": "[依赖库] 压缩算法",
    "certifi": "[依赖库] SSL 证书验证",
    "chardet": "[依赖库] 字符编码检测",
    "charset-normalizer": "[依赖库] 字符编码标准化",
    "et_xmlfile": "[依赖库] openpyxl XML 处理",
    "idna": "[依赖库] 国际化域名支持",
    "inflate64": "[依赖库] py7zr 解压算法",
    "multivolumefile": "[依赖库] py7zr 多卷文件处理",
    "packaging": "[依赖库] 包版本号处理",
    "pefile": "[依赖库] pyinstaller EXE 分析",
    "pybcj": "[依赖库] py7zr BCJ 过滤器",
    "pycryptodomex": "[依赖库] py7zr 加密解密",
    "pyinstaller-hooks-contrib": "[依赖库] pyinstaller 第三方钩子",
    "pyppmd": "[依赖库] py7zr PPMd 算法",
    "PyQt5-Qt5": "[依赖库] PyQt5 Qt 核心库",
    "PyQt5_sip": "[依赖库] PyQt5 C++/Python 绑定",
    "PySide6_Addons": "[依赖库] PySide6 附加组件",
    "PySide6_Essentials": "[依赖库] PySide6 核心组件",
    "shiboken6": "[依赖库] PySide6 C++/Python 绑定",
    "pywin32-ctypes": "[依赖库] pyinstaller Windows API",
    "setuptools": "[依赖库] Python 打包工具",
    "soupsieve": "[依赖库] beautifulsoup4 CSS 选择器",
    "typing_extensions": "[依赖库] Python 类型扩展",
    "urllib3": "[依赖库] requests HTTP 连接池",
    "pyqt6": "桌面软件界面开发（PyQt6 主库）",
    "PyQt6-Qt6": "[依赖库] PyQt6 Qt 核心库",
    "PyQt6-sip": "[依赖库] PyQt6 C++/Python 绑定",
    "PyQt6-Charts": "PyQt6 图表控件",
    "PyQt6-Charts-Qt6": "[依赖库] PyQt6 图表底层库",
    "nvidia-ml-py": "读取 N 卡信息（显存/温度）",
}

_LIB_DESC_LOWER = {k.lower(): v for k, v in LIB_DESC.items()}

COMMON_LIB_NAMES = [
    "openpyxl",
    "requests",
    "beautifulsoup4",
    "PyQt5",
    "PySide6",
    "pyinstaller",
    "Nuitka",
    "py7zr",
    "psutil",
    "colorama",
    "polib",
    "texttable",
    "pandas",
    "numpy",
    "matplotlib",
    "lxml",
    "selenium",
    "aiohttp",
    "pyperclip",
    "pillow",
    "opencv-python",
    "tqdm",
    "schedule",
    "pyyaml",
    "python-dotenv",
    "click",
    "Flask",
    "Django",
    "jupyter",
    "pytest",
    "black",
    "xlrd",
    "xlwt",
    "pywin32",
]


def get_lib_description(name):
    desc = LIB_DESC.get(name)
    if desc:
        return desc
    desc = _LIB_DESC_LOWER.get(name.lower())
    if desc:
        return desc
    return "第三方库"


def get_installed_libs():
    try:
        result = subprocess.run(
            PYTHON_CMD + ["-m", "pip", "list", "--format=freeze"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=CREATE_NO_WINDOW,
            env=_get_sub_env(),
            timeout=120,
        )
        if result.returncode != 0:
            return []
        libs = []
        for line in result.stdout.strip().splitlines():
            if "==" in line:
                name, ver = line.split("==", 1)
                libs.append((name.strip(), ver.strip()))
        return libs
    except Exception:
        return []


def is_lib_installed(name):
    for lib_name, _ in get_installed_libs():
        if lib_name.lower() == name.lower():
            return True
    return False


def run_command(cmd_list, callback=None, timeout=None):
    try:
        process = subprocess.Popen(
            cmd_list,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=CREATE_NO_WINDOW,
            env=_get_sub_env(),
        )
    except Exception as e:
        return str(e), -1

    output_lines = []
    timed_out = [False]

    def reader():
        try:
            for line in iter(process.stdout.readline, ""):
                output_lines.append(line)
                if callback:
                    callback(line.rstrip())
        except Exception:
            pass

    reader_thread = threading.Thread(target=reader, daemon=True)
    reader_thread.start()

    if timeout:
        try:
            process.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            timed_out[0] = True
            try:
                process.kill()
            except Exception:
                pass
            try:
                process.wait(timeout=5)
            except Exception:
                pass
    else:
        process.wait()

    reader_thread.join(timeout=3)
    try:
        process.stdout.close()
    except Exception:
        pass

    if timed_out[0]:
        return "".join(output_lines), -2
    return "".join(output_lines), process.returncode


class PyAssistantApp:
    def __init__(self, root):
        self.root = root
        self.build_date = "2026-10-03 00:00"
        root.title(f"{APP_NAME} {VERSION}")
        win_w, win_h = 850, 600
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        pos_x = (screen_w - win_w) // 2
        pos_y = (screen_h - win_h) // 2
        root.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        root.minsize(850, 600)

        self.script_path = tk.StringVar()
        self.icon_path = tk.StringVar()
        self.n_icon_path = tk.StringVar()
        self.opt_onefile = tk.BooleanVar(value=True)
        self.opt_windowed = tk.BooleanVar(value=False)
        self.opt_icon = tk.BooleanVar(value=False)
        self.opt_n_onefile = tk.BooleanVar(value=True)
        self.opt_n_windowed = tk.BooleanVar(value=False)
        self.opt_n_icon = tk.BooleanVar(value=False)
        self.control_buttons = []
        self.py_update_has_update = False

        self.busy = False
        self._last_env_ok = False

        self._env_cache = None
        self._env_cache_time = 0

        self.split_y = 0
        self.split_bar_height = 8
        self.split_min_top = 250
        self.split_min_bottom = 120
        self.split_report_height = 240
        self.split_user_dragged = False
        self.split_dragging = False
        self.split_drag_start_y = 0
        self.split_drag_start_split = 0

        self._create_widgets()
        root.deiconify()
        self._refresh_status(force=True)
        self.root.after(150, self._set_initial_split)

        self._log("=" * 60)
        self._log("🔍 启动环境检测")
        self._log("-" * 60)
        py_ver, pip_ver, env_ok = self._get_env_status(force=True)
        if py_ver:
            self._log(f"✅ Python 版本: {py_ver}")
        else:
            self._log("❌ Python 未安装")
        if pip_ver:
            self._log(f"✅ pip 版本: {pip_ver}")
        else:
            self._log("❌ pip 未安装")
        if env_ok:
            self._log("💡 环境正常，可以正常使用。")
        else:
            self._log("⚠️ 当前系统未检测到 Python 环境。")
            self._log("请点击「安装区」→「安装 Python」查看安装指引。")
            self._log("安装完成后请重新打开本工具。")
        self._log("=" * 60)
        self._log("")


    def _get_python_version(self):
        try:
            result = subprocess.run(
                PYTHON_CMD + ["--version"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=CREATE_NO_WINDOW,
                env=_get_sub_env(),
                timeout=30,
            )
            if result.returncode == 0:
                version = result.stdout.strip() or result.stderr.strip()
                if version:
                    return version.replace("Python ", "")
            return None
        except Exception:
            return None

    def _get_pip_version(self):
        try:
            result = subprocess.run(
                PYTHON_CMD + ["-m", "pip", "--version"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=CREATE_NO_WINDOW,
                env=_get_sub_env(),
                timeout=30,
            )
            if result.returncode == 0:
                parts = result.stdout.strip().split()
                if len(parts) >= 2:
                    return parts[1]
            return None
        except Exception:
            return None

    def _get_env_status(self, force=False):
        now = time.time()
        if (not force) and self._env_cache is not None and (now - self._env_cache_time) < 3:
            return self._env_cache
        py_ver = self._get_python_version()
        pip_ver = self._get_pip_version()
        ok = (py_ver is not None and pip_ver is not None)
        self._env_cache = (py_ver, pip_ver, ok)
        self._env_cache_time = time.time()
        return self._env_cache

    def _refresh_status(self, force=False):
        py_ver, pip_ver, env_ok = self._get_env_status(force=force)
        self._last_env_ok = env_ok
        if py_ver:
            self.py_status.config(text=f"✅ Python {py_ver}", fg="green")
        else:
            self.py_status.config(text="❌ Python 未安装", fg="red")
        if pip_ver:
            self.pip_status.config(text=f"✅ pip {pip_ver}", fg="green")
        else:
            self.pip_status.config(text="❌ pip 未安装", fg="red")
        if env_ok:
            self.env_status.config(text="🟢 环境正常", fg="green")
        else:
            self.env_status.config(text="🔴 环境异常", fg="red")
        self._update_buttons(env_ok)

    def _version_tuple(self, v):
        nums = []
        for part in str(v).split("."):
            digits = ""
            for ch in part:
                if ch.isdigit():
                    digits += ch
                else:
                    break
            if digits:
                nums.append(int(digits))
        return tuple(nums)

    def _get_latest_python_version(self):
        import json
        import urllib.request
        req = urllib.request.Request(
            "https://endoflife.date/api/python.json",
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"},
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        candidates = []
        for item in data:
            if item.get("development") or item.get("discontinued"):
                continue
            cycle = str(item.get("cycle", ""))
            latest = str(item.get("latest", ""))
            if not cycle or not latest:
                continue
            candidates.append((self._version_tuple(cycle), latest))
        if not candidates:
            return None
        candidates.sort(key=lambda x: x[0], reverse=True)
        return candidates[0][1]

    def _check_python_update(self):
        self.py_update_has_update = False
        self.py_update_status.config(text="⏳ 检测Python更新中...", fg="green", cursor="hand2")

        def work():
            try:
                latest = self._get_latest_python_version()
                current = self._get_python_version()
            except Exception:
                latest = None
                current = None

            def done():
                try:
                    if latest and current:
                        if self._version_tuple(latest) > self._version_tuple(current):
                            self.py_update_has_update = True
                            self.py_update_status.config(text=f"⚠️ Python有新版本 {latest}", fg="red")
                        else:
                            self.py_update_status.config(text="✅ Python已是最新", fg="green")
                    else:
                        self.py_update_status.config(text="⚠️ 检测失败,点击重试", fg="#666666")
                except Exception:
                    self.py_update_status.config(text="⚠️ 检测失败,点击重试", fg="#666666")

            self.root.after(0, done)

        threading.Thread(target=work, daemon=True).start()

    def _on_py_update_click(self, event=None):
        if self.py_update_has_update:
            webbrowser.open("https://www.python.org/downloads/")
        else:
            self._check_python_update()

    def _set_busy(self, busy):
        self.busy = busy
        self._update_buttons(self._last_env_ok)

    def _update_buttons(self, env_ok):
        env_state = "normal" if (env_ok and not self.busy) else "disabled"
        idle_state = "disabled" if self.busy else "normal"

        for btn in self.control_buttons:
            btn.config(state=env_state)

        if hasattr(self, "btn_check_env"):
            self.btn_check_env.config(state=idle_state)
        if hasattr(self, "btn_refresh_status"):
            self.btn_refresh_status.config(state=idle_state)
        if hasattr(self, "btn_install_py"):
            self.btn_install_py.config(state=idle_state)

        if hasattr(self, "btn_list_libs"):
            self.btn_list_libs.config(state=env_state)
        if hasattr(self, "btn_check_lib_update"):
            self.btn_check_lib_update.config(state=env_state)

        icon_ok = self.opt_icon.get() and env_ok and not self.busy
        if hasattr(self, "icon_entry"):
            self.icon_entry.config(state="normal" if icon_ok else "disabled")
        if hasattr(self, "icon_btn"):
            self.icon_btn.config(state="normal" if icon_ok else "disabled")

        if hasattr(self, "script_entry"):
            self.script_entry.config(state=env_state)
        if hasattr(self, "opt_onefile_btn"):
            self.opt_onefile_btn.config(state=env_state)
        if hasattr(self, "opt_windowed_btn"):
            self.opt_windowed_btn.config(state=env_state)
        if hasattr(self, "opt_icon_btn"):
            self.opt_icon_btn.config(state=env_state)

        n_icon_ok = self.opt_n_icon.get() and env_ok and not self.busy
        if hasattr(self, "opt_n_onefile_btn"):
            self.opt_n_onefile_btn.config(state=env_state)
        if hasattr(self, "opt_n_windowed_btn"):
            self.opt_n_windowed_btn.config(state=env_state)
        if hasattr(self, "opt_n_icon_btn"):
            self.opt_n_icon_btn.config(state=env_state)
        if hasattr(self, "n_icon_entry"):
            self.n_icon_entry.config(state="normal" if n_icon_ok else "disabled")
        if hasattr(self, "n_icon_btn"):
            self.n_icon_btn.config(state="normal" if n_icon_ok else "disabled")
        if hasattr(self, "n_script_entry"):
            self.n_script_entry.config(state=env_state)


    def _center_window(self, parent, child, width, height):
        try:
            parent.update_idletasks()
        except Exception:
            pass
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        x = parent_x + (parent_w - width) // 2
        y = parent_y + (parent_h - height) // 2
        child.geometry(f"{width}x{height}+{x}+{y}")

    def _set_initial_split(self):
        try:
            total_height = self.split_container.winfo_height()
            if total_height <= 0:
                return
            self.split_y = total_height - self.split_bar_height - self.split_report_height
            self._clamp_split()
            self._apply_split()
        except Exception:
            pass

    def _clamp_split(self):
        total_height = self.split_container.winfo_height()
        if total_height <= 0:
            return
        min_y = self.split_min_top
        max_y = total_height - self.split_bar_height - self.split_min_bottom
        if max_y < min_y:
            max_y = min_y
        self.split_y = max(min_y, min(self.split_y, max_y))

    def _apply_split(self):
        if not hasattr(self, "split_container"):
            return
        total_height = self.split_container.winfo_height()
        if total_height <= 0:
            return
        self._clamp_split()
        top_height = self.split_y
        bottom_y = self.split_y + self.split_bar_height
        bottom_height = total_height - bottom_y
        if bottom_height < 0:
            bottom_height = 0
        self.top_frame.place(x=0, y=0, relwidth=1, height=top_height)
        self.split_separator_frame.place(x=0, y=self.split_y, relwidth=1, height=self.split_bar_height)
        self.bottom_frame.place(x=0, y=bottom_y, relwidth=1, height=bottom_height)

    def _on_split_container_configure(self, event=None):
        if self.split_dragging:
            return
        total_height = self.split_container.winfo_height()
        if total_height <= 0:
            return
        if not self.split_user_dragged:
            self.split_y = total_height - self.split_bar_height - self.split_report_height
        self._clamp_split()
        self._apply_split()

    def _split_button_press(self, event):
        self.split_dragging = True
        self.split_user_dragged = True
        self.split_drag_start_y = event.y_root
        self.split_drag_start_split = self.split_y
        self.split_separator_frame.configure(cursor="sb_v_double_arrow")

    def _split_motion(self, event):
        if not self.split_dragging:
            return
        delta = event.y_root - self.split_drag_start_y
        new_y = self.split_drag_start_split + delta
        total_height = self.split_container.winfo_height()
        min_y = self.split_min_top
        max_y = total_height - self.split_bar_height - self.split_min_bottom
        if max_y < min_y:
            max_y = min_y
        new_y = max(min_y, min(new_y, max_y))
        bottom_y = new_y + self.split_bar_height
        bottom_height = total_height - bottom_y
        if bottom_height < 0:
            bottom_height = 0
        self.top_frame.place(x=0, y=0, relwidth=1, height=new_y)
        self.split_separator_frame.place(x=0, y=new_y, relwidth=1, height=self.split_bar_height)
        self.bottom_frame.place(x=0, y=bottom_y, relwidth=1, height=bottom_height)

    def _split_button_release(self, event):
        if not self.split_dragging:
            return
        total_height = self.split_container.winfo_height()
        min_y = self.split_min_top
        max_y = total_height - self.split_bar_height - self.split_min_bottom
        delta = event.y_root - self.split_drag_start_y
        new_y = self.split_drag_start_split + delta
        if max_y < min_y:
            max_y = min_y
        new_y = max(min_y, min(new_y, max_y))
        self.split_y = new_y
        self.split_dragging = False
        self.split_separator_frame.configure(cursor="sb_v_double_arrow")
        self._apply_split()


    def _create_widgets(self):
        self.status_frame = tk.Frame(self.root, bg="#f0f0f0", height=32)
        self.status_frame.pack(fill=tk.X, padx=0, pady=0)
        self.status_frame.pack_propagate(False)

        self.py_status = tk.Label(self.status_frame, text="Python: 检测中...", font=("微软雅黑", 10), fg="green", bg="#f0f0f0")
        self.py_status.pack(side=tk.LEFT, padx=10, pady=4)

        self.pip_status = tk.Label(self.status_frame, text="pip: 检测中...", font=("微软雅黑", 10), fg="green", bg="#f0f0f0")
        self.pip_status.pack(side=tk.LEFT, padx=10, pady=4)

        self.py_update_status = tk.Label(self.status_frame, text="🔍 点击检测Python新版本", font=("微软雅黑", 10), fg="#0066cc", bg="#f0f0f0", cursor="hand2")
        self.py_update_status.pack(side=tk.LEFT, padx=10, pady=4)
        self.py_update_status.bind("<Button-1>", self._on_py_update_click)

        self.env_status = tk.Label(self.status_frame, text="🟢 环境正常", font=("微软雅黑", 9), fg="green", bg="#f0f0f0")
        self.env_status.pack(side=tk.RIGHT, padx=10, pady=4)

        sep_line = tk.Frame(self.root, height=2, bg="#cccccc")
        sep_line.pack(fill=tk.X, padx=0, pady=0)

        self.split_container = tk.Frame(self.root, bg="white")
        self.split_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.split_container.bind("<Configure>", self._on_split_container_configure)

        self.top_frame = tk.Frame(self.split_container, bg="white")
        self.bottom_frame = tk.Frame(self.split_container, bg="white")

        self.split_separator_frame = tk.Frame(self.split_container, bg="#d0d0d0", cursor="sb_v_double_arrow")
        self.split_highlight = tk.Frame(self.split_separator_frame, bg="#f7f7f7", height=1)
        self.split_highlight.place(x=0, y=0, relwidth=1, height=1)
        self.split_middle = tk.Frame(self.split_separator_frame, bg="#dddddd")
        self.split_middle.place(x=0, y=1, relwidth=1, relheight=1, height=-2)
        self.split_shadow = tk.Frame(self.split_separator_frame, bg="#c2c2c2", height=1)
        self.split_shadow.place(x=0, y=self.split_bar_height - 1, relwidth=1, height=1)

        for widget in (self.split_separator_frame, self.split_highlight, self.split_middle, self.split_shadow):
            widget.bind("<ButtonPress-1>", self._split_button_press)
            widget.bind("<B1-Motion>", self._split_motion)
            widget.bind("<ButtonRelease-1>", self._split_button_release)

        self.top_frame.columnconfigure(0, weight=1, uniform="main_columns")
        self.top_frame.columnconfigure(2, weight=1, uniform="main_columns")
        self.top_frame.columnconfigure(4, weight=1, uniform="main_columns")
        self.top_frame.columnconfigure(1, weight=0)
        self.top_frame.columnconfigure(3, weight=0)
        self.top_frame.rowconfigure(0, weight=1)

        self._create_check_panel(self.top_frame, 0)
        sep1 = ttk.Separator(self.top_frame, orient=tk.VERTICAL)
        sep1.grid(row=0, column=1, sticky="ns", padx=3, pady=8)

        self._create_install_panel(self.top_frame, 2)
        sep2 = ttk.Separator(self.top_frame, orient=tk.VERTICAL)
        sep2.grid(row=0, column=3, sticky="ns", padx=3, pady=8)

        self._create_pack_panel(self.top_frame, 4)

        report_label_frame = tk.Frame(self.bottom_frame, bg="white")
        report_label_frame.pack(fill=tk.X, pady=(0, 2))
        tk.Label(report_label_frame, text="📋 报告区", font=("微软雅黑", 10, "bold"), bg="white").pack(side=tk.LEFT)
        tk.Button(report_label_frame, text="清空", command=self._clear_output, font=("微软雅黑", 9)).pack(side=tk.RIGHT)

        self.output_text = scrolledtext.ScrolledText(self.bottom_frame, font=("Consolas", 9), wrap=tk.WORD, bg="#fafafa")
        self.output_text.pack(fill=tk.BOTH, expand=True)

        self._sym_tags = {}
        for _i, (_sym, _color) in enumerate(REPORT_SYMBOL_COLORS.items()):
            _tag = "sym_%d" % _i
            self.output_text.tag_config(_tag, foreground=_color)
            self._sym_tags[_sym] = _tag

        self.output_menu = tk.Menu(self.root, tearoff=0)
        self.output_menu.add_command(label="复制", command=self._copy_output_selection)
        self.output_menu.add_command(label="全选", command=self._select_all_output)
        self.output_menu.add_separator()
        self.output_menu.add_command(label="清空", command=self._clear_output)
        self.output_text.bind("<Button-3>", self._show_output_menu)

        bottom_frame = tk.Frame(self.root, bg="#f0f0f0", height=28)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM)
        bottom_frame.pack_propagate(False)
        tk.Label(bottom_frame, text=f"作者：{AUTHOR}  |  版本：{VERSION}  |  {self.build_date}", font=("微软雅黑", 9), fg="#666666", bg="#f0f0f0").pack(side=tk.LEFT, padx=10)
        tk.Button(bottom_frame, text="关于", font=("微软雅黑", 9), command=self._show_about, relief=tk.FLAT, bg="#f0f0f0", fg="#0066cc", cursor="hand2").pack(side=tk.RIGHT, padx=10)

    def _create_check_panel(self, parent, col):
        panel = tk.Frame(parent, relief=tk.GROOVE, bd=2)
        panel.grid(row=0, column=col, sticky="nsew", padx=5, pady=5)
        tk.Label(panel, text="🔍 检查区", font=("微软雅黑", 12, "bold")).pack(anchor="w", padx=10, pady=8)

        self.btn_check_env = tk.Button(panel, text="检测环境", font=("微软雅黑", 10), width=14, command=self._check_env)
        self.btn_check_env.pack(anchor="center", padx=10, pady=4)

        self.btn_refresh_status = tk.Button(panel, text="刷新状态", font=("微软雅黑", 10), width=14, command=self._refresh_manual)
        self.btn_refresh_status.pack(anchor="center", padx=10, pady=4)

        self.btn_list_libs = tk.Button(panel, text="查看所有库", font=("微软雅黑", 10), width=14, command=self._view_all_libs)
        self.btn_list_libs.pack(anchor="center", padx=10, pady=4)

        self.btn_export = tk.Button(panel, text="导出到清单", font=("微软雅黑", 10), width=14, command=self._export_requirements)
        self.btn_export.pack(anchor="center", padx=10, pady=4)

        self.control_buttons.extend([self.btn_export])

    def _create_install_panel(self, parent, col):
        panel = tk.Frame(parent, relief=tk.GROOVE, bd=2)
        panel.grid(row=0, column=col, sticky="nsew", padx=5, pady=5)
        tk.Label(panel, text="📦 安装区", font=("微软雅黑", 12, "bold")).pack(anchor="w", padx=10, pady=8)

        self.btn_install_py = tk.Button(panel, text="安装 Python", font=("微软雅黑", 10), width=14, command=self._install_python_guide)
        self.btn_install_py.pack(anchor="center", padx=10, pady=4)

        self.btn_restore = tk.Button(panel, text="从清单恢复库", font=("微软雅黑", 10), width=14, command=self._restore_from_file)
        self.btn_restore.pack(anchor="center", padx=10, pady=4)

        self.btn_common = tk.Button(panel, text="安装常用库", font=("微软雅黑", 10), width=14, command=self._install_common_libs)
        self.btn_common.pack(anchor="center", padx=10, pady=4)

        self.btn_check_lib_update = tk.Button(panel, text="检查库更新", font=("微软雅黑", 10), width=14, command=self._check_lib_update)
        self.btn_check_lib_update.pack(anchor="center", padx=10, pady=4)

        self.control_buttons.extend([self.btn_restore, self.btn_common])

    def _create_pack_panel(self, parent, col):
        panel = tk.Frame(parent, relief=tk.GROOVE, bd=2)
        panel.grid(row=0, column=col, sticky="nsew", padx=5, pady=5)

        header = tk.Frame(panel)
        header.pack(anchor="w", padx=10, pady=(8, 2), fill=tk.X)
        tk.Label(header, text="🚀 打包区", font=("微软雅黑", 12, "bold")).pack(side=tk.LEFT)
        self.btn_pack = tk.Button(header, text="开始打包", font=("微软雅黑", 10, "bold"), bg="#4CAF50", fg="white", width=14, command=self._start_pack)
        self.btn_pack.pack(side=tk.RIGHT)

        self.notebook_pack = ttk.Notebook(panel)
        self.notebook_pack.pack(fill=tk.BOTH, expand=True, padx=5, pady=(2, 5))
        self.notebook_pack.bind("<<NotebookTabChanged>>", self._on_pack_tab_changed)

        self.pack_tab_pyinstaller = tk.Frame(self.notebook_pack)
        self.pack_tab_nuitka = tk.Frame(self.notebook_pack)
        self.notebook_pack.add(self.pack_tab_pyinstaller, text="PyInstaller")
        self.notebook_pack.add(self.pack_tab_nuitka, text="Nuitka")

        tab = self.pack_tab_pyinstaller
        script_frame = tk.Frame(tab)
        script_frame.pack(anchor="w", padx=10, pady=4, fill=tk.X)
        tk.Label(script_frame, text="脚本:", font=("微软雅黑", 9)).pack(side=tk.LEFT)
        self.script_entry = tk.Entry(script_frame, textvariable=self.script_path, font=("微软雅黑", 9), width=20)
        self.script_entry.pack(side=tk.LEFT, padx=4, fill=tk.X, expand=True)
        self.btn_select_script = tk.Button(script_frame, text="...", width=3, command=self._select_script)
        self.btn_select_script.pack(side=tk.RIGHT)

        self.opt_onefile_btn = tk.Checkbutton(tab, text="单文件打包 (--onefile)", variable=self.opt_onefile, font=("微软雅黑", 9), anchor="w")
        self.opt_onefile_btn.pack(anchor="w", padx=12, pady=2)

        self.opt_windowed_btn = tk.Checkbutton(tab, text="无控制台窗口 (--windowed)", variable=self.opt_windowed, font=("微软雅黑", 9), anchor="w")
        self.opt_windowed_btn.pack(anchor="w", padx=12, pady=2)

        icon_frame = tk.Frame(tab)
        icon_frame.pack(anchor="w", padx=12, pady=2, fill=tk.X)
        self.opt_icon_btn = tk.Checkbutton(icon_frame, text="自定义图标", variable=self.opt_icon, font=("微软雅黑", 9), anchor="w", command=self._toggle_icon)
        self.opt_icon_btn.pack(side=tk.LEFT)
        self.icon_entry = tk.Entry(icon_frame, textvariable=self.icon_path, font=("微软雅黑", 9), width=12, state="disabled")
        self.icon_entry.pack(side=tk.LEFT, padx=4, fill=tk.X, expand=True)
        self.icon_btn = tk.Button(icon_frame, text="...", width=3, command=lambda: self._select_icon(self.icon_path), state="disabled")
        self.icon_btn.pack(side=tk.RIGHT)

        self.control_buttons.extend([self.btn_select_script, self.btn_pack])

        tab = self.pack_tab_nuitka
        script_frame = tk.Frame(tab)
        script_frame.pack(anchor="w", padx=10, pady=4, fill=tk.X)
        tk.Label(script_frame, text="脚本:", font=("微软雅黑", 9)).pack(side=tk.LEFT)
        self.n_script_entry = tk.Entry(script_frame, textvariable=self.script_path, font=("微软雅黑", 9), width=20)
        self.n_script_entry.pack(side=tk.LEFT, padx=4, fill=tk.X, expand=True)
        self.btn_n_select_script = tk.Button(script_frame, text="...", width=3, command=self._select_script)
        self.btn_n_select_script.pack(side=tk.RIGHT)

        self.opt_n_onefile_btn = tk.Checkbutton(tab, text="单文件打包 (--onefile)", variable=self.opt_n_onefile, font=("微软雅黑", 9), anchor="w")
        self.opt_n_onefile_btn.pack(anchor="w", padx=12, pady=2)

        self.opt_n_windowed_btn = tk.Checkbutton(tab, text="无控制台窗口", variable=self.opt_n_windowed, font=("微软雅黑", 9), anchor="w")
        self.opt_n_windowed_btn.pack(anchor="w", padx=12, pady=2)

        icon_frame = tk.Frame(tab)
        icon_frame.pack(anchor="w", padx=12, pady=2, fill=tk.X)
        self.opt_n_icon_btn = tk.Checkbutton(icon_frame, text="自定义图标", variable=self.opt_n_icon, font=("微软雅黑", 9), anchor="w", command=self._toggle_n_icon)
        self.opt_n_icon_btn.pack(side=tk.LEFT)
        self.n_icon_entry = tk.Entry(icon_frame, textvariable=self.n_icon_path, font=("微软雅黑", 9), width=12, state="disabled")
        self.n_icon_entry.pack(side=tk.LEFT, padx=4, fill=tk.X, expand=True)
        self.n_icon_btn = tk.Button(icon_frame, text="...", width=3, command=lambda: self._select_icon(self.n_icon_path), state="disabled")
        self.n_icon_btn.pack(side=tk.RIGHT)

        env_btn_row = tk.Frame(tab)
        env_btn_row.pack(anchor="w", padx=10, pady=4)
        self.btn_n_check_env = tk.Button(env_btn_row, text="检测Nuitka环境", font=("微软雅黑", 9), width=14, command=self._check_nuitka_env)
        self.btn_n_check_env.pack(side=tk.LEFT)
        self.btn_n_clear_env = tk.Button(env_btn_row, text="清除Nuitka环境", font=("微软雅黑", 9), width=14, command=self._clear_nuitka_env)
        self.btn_n_clear_env.pack(side=tk.LEFT, padx=(8, 0))

        self.control_buttons.extend([self.btn_n_select_script, self.btn_n_check_env, self.btn_n_clear_env])

    def _on_pack_tab_changed(self, event=None):
        try:
            current = self.notebook_pack.select()
            if current == str(self.pack_tab_nuitka):
                self.btn_pack.config(command=self._start_pack_nuitka)
            else:
                self.btn_pack.config(command=self._start_pack)
        except Exception:
            pass

    def _log(self, text):
        self.root.after(0, self._log_ui, text)

    def _log_ui(self, text):
        try:
            start = self.output_text.index("end-1c")
            self.output_text.insert(tk.END, text + "\n")
            stop = self.output_text.index("end-1c")
            sym_tags = getattr(self, "_sym_tags", None)
            if sym_tags:
                cnt = tk.IntVar()
                for sym, tag in sym_tags.items():
                    key = sym.rstrip("\ufe0f")
                    idx = start
                    while True:
                        try:
                            pos = self.output_text.search(key, idx, stopindex=stop, count=cnt)
                            n = cnt.get() if pos else 0
                        except Exception:
                            pos = self.output_text.search(key, idx, stopindex=stop)
                            n = 0
                        if not pos:
                            break
                        if n <= 0:
                            n = sum(2 if ord(c) > 0xFFFF else 1 for c in key)
                        end = "%s+%dc" % (pos, n)
                        self.output_text.tag_add(tag, pos, end)
                        idx = end
            self.output_text.see(tk.END)
            self.root.update_idletasks()
        except Exception:
            pass

    def _clear_output(self):
        self.output_text.delete(1.0, tk.END)

    def _show_output_menu(self, event):
        try:
            has_sel = bool(self.output_text.tag_ranges(tk.SEL))
            self.output_menu.entryconfig("复制", state=(tk.NORMAL if has_sel else tk.DISABLED))
        except Exception:
            pass
        try:
            self.output_menu.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                self.output_menu.grab_release()
            except Exception:
                pass

    def _copy_output_selection(self):
        try:
            content = self.output_text.get(tk.SEL_FIRST, tk.SEL_LAST)
        except Exception:
            return
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(content)
        except Exception:
            pass

    def _select_all_output(self):
        try:
            self.output_text.tag_add(tk.SEL, "1.0", tk.END)
            self.output_text.mark_set(tk.INSERT, "1.0")
            self.output_text.see("1.0")
        except Exception:
            pass

    def _open_folder(self, path):
        try:
            if not path or not os.path.isdir(path):
                return
            abs_path = os.path.abspath(path)
            if sys.platform == "win32":
                os.startfile(abs_path)
            else:
                webbrowser.open("file://" + abs_path)
        except Exception:
            pass


    def _check_env(self):
        self._log("=" * 60)
        self._log("🔍 环境检测报告")
        self._log("-" * 60)
        py_ver, pip_ver, env_ok = self._get_env_status(force=True)
        if py_ver:
            self._log(f"✅ Python 版本: {py_ver}")
        else:
            self._log("❌ Python 未安装")
        if pip_ver:
            self._log(f"✅ pip 版本: {pip_ver}")
        else:
            self._log("❌ pip 未安装")
        import platform
        self._log(f"✅ 系统: {platform.system()} {platform.release()} ({platform.machine()})")

        try:
            result = subprocess.run(
                PYTHON_CMD + ["-m", "pip", "config", "list"],
                capture_output=True, text=True,
                encoding="utf-8", errors="replace",
                creationflags=CREATE_NO_WINDOW, env=_get_sub_env(),
                timeout=30,
            )
            found = False
            for line in (result.stdout or "").splitlines():
                if "global.index-url" in line:
                    self._log(f"✅ pip 源: {line.split('=')[-1].strip()}")
                    found = True
                    break
            if not found:
                self._log("✅ pip 源: https://pypi.org/simple (默认)")
        except Exception:
            self._log("✅ pip 源: https://pypi.org/simple (默认)")

        libs = get_installed_libs()
        self._log(f"✅ 已安装第三方库: {len(libs)} 个")
        self._log("")
        self._log("📦 已安装的常用库：")
        installed_names = {name.lower() for name, _ in libs}
        common_installed = []
        for name in COMMON_LIB_NAMES:
            if name.lower() in installed_names:
                desc = get_lib_description(name)
                common_installed.append(f"  - {name:20} ({desc})")
        if common_installed:
            for item in common_installed:
                self._log(item)
        else:
            self._log("  （暂无常用库）")
        self._log("")
        if env_ok:
            self._log("💡 环境正常，可以正常使用。")
        else:
            self._log("⚠️ 环境异常，请检查 Python 安装。")
            self._log("请点击「安装区」→「安装 Python」查看安装指引。")
        self._log("=" * 60)
        self._log("")
        self._refresh_status(force=True)

    def _refresh_manual(self):
        self._log("🔄 正在刷新环境状态...")
        self._refresh_status(force=True)
        _, _, env_ok = self._get_env_status()
        if env_ok:
            self._log("✅ 状态已刷新，环境正常。")
        else:
            self._log("⚠️ 状态已刷新，环境仍异常。请检查 Python 安装。")
        self._log("")


    def _libs_sort(self, libs):
        pip_item = None
        others = []
        for name, ver in libs:
            if name.lower() == "pip":
                pip_item = (name, ver)
            else:
                others.append((name, ver))
        others.sort(key=lambda x: x[0].lower())
        result = []
        if pip_item:
            result.append(pip_item)
        result.extend(others)
        return result

    def _print_all_libs_to_report(self, libs):
        sorted_libs = self._libs_sort(libs)
        self._log("=" * 60)
        self._log("📦 已安装库列表")
        self._log("-" * 60)
        self._log(f"共 {len(libs)} 个库：")
        self._log("")
        for name, ver in sorted_libs:
            desc = get_lib_description(name)
            self._log(f"  {name}=={ver}    # {desc}")
        self._log("")
        self._log("=" * 60)
        self._log("")

    def _make_libs_text(self, libs):
        sorted_libs = self._libs_sort(libs)
        lines = [f"共 {len(libs)} 个库：", ""]
        for name, ver in sorted_libs:
            desc = get_lib_description(name)
            lines.append(f"{name}=={ver}    # {desc}")
        return "\n".join(lines)

    def _view_all_libs(self):
        _, _, env_ok = self._get_env_status()
        if not env_ok:
            messagebox.showwarning("无法查看", "⚠️ 当前系统未检测到 Python 环境。\n\n请点击「安装 Python」进行安装，安装完成后重新打开本工具。")
            return
        libs = get_installed_libs()
        if not libs:
            messagebox.showwarning("提示", "没有检测到已安装的库。")
            return

        self._print_all_libs_to_report(libs)

        dialog = tk.Toplevel(self.root)
        dialog.title("📦 查看所有库")
        self._center_window(self.root, dialog, 620, 520)
        dialog.transient(self.root)
        dialog.grab_set()

        top_frame = tk.Frame(dialog)
        top_frame.pack(fill=tk.X, padx=10, pady=5)
        dialog.count_label = tk.Label(top_frame, text=f"共 {len(libs)} 个已安装的库：", font=("微软雅黑", 10, "bold"))
        dialog.count_label.pack(side=tk.LEFT)
        tk.Button(top_frame, text="全选", command=lambda: self._all_libs_select(dialog, True)).pack(side=tk.RIGHT, padx=2)
        tk.Button(top_frame, text="反选", command=lambda: self._all_libs_select(dialog, False)).pack(side=tk.RIGHT, padx=2)

        list_frame = tk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        canvas = tk.Canvas(list_frame, highlightthickness=0)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=canvas.yview)
        scroll_frame = tk.Frame(canvas)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas_frame = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        scroll_frame.bind("<Configure>", on_frame_configure)

        def on_mousewheel(event):
            try:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass

        def bind_wheel(e=None):
            canvas.bind_all("<MouseWheel>", on_mousewheel)

        def unbind_wheel(e=None):
            try:
                canvas.unbind_all("<MouseWheel>")
            except Exception:
                pass

        canvas.bind("<Enter>", bind_wheel)
        canvas.bind("<Leave>", unbind_wheel)

        def on_destroy(e):
            if e.widget is dialog:
                unbind_wheel()
        dialog.bind("<Destroy>", on_destroy)

        dialog.vars_map = {}
        dialog.row_refs = {}
        dialog.scroll_frame = scroll_frame
        dialog.canvas = canvas
        dialog.libs_data = list(libs)

        self._build_all_libs_rows(scroll_frame, dialog, libs)

        bottom_frame = tk.Frame(dialog)
        bottom_frame.pack(fill=tk.X, padx=10, pady=10)

        btn_container = tk.Frame(bottom_frame)
        btn_container.pack()
        tk.Button(btn_container, text="批量卸载", font=("微软雅黑", 10), bg="#cc3333", fg="white", command=lambda: self._batch_uninstall(dialog)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_container, text="复制清单", font=("微软雅黑", 10), command=lambda: self._copy_libs_to_clipboard(dialog)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_container, text="刷新", font=("微软雅黑", 10), command=lambda: self._refresh_all_libs(dialog)).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_container, text="关闭", font=("微软雅黑", 10), command=dialog.destroy).pack(side=tk.LEFT, padx=5)

        dialog._resize_job = None

        def apply_canvas_width(w):
            dialog._resize_job = None
            try:
                canvas.itemconfig(canvas_frame, width=w)
            except Exception:
                pass

        def on_resize(event):
            w = max(event.width - 10, 1)
            if dialog._resize_job is not None:
                try:
                    dialog.after_cancel(dialog._resize_job)
                except Exception:
                    pass
            dialog._resize_job = dialog.after(40, lambda: apply_canvas_width(w))
        canvas.bind("<Configure>", on_resize)

    def _copy_libs_to_clipboard(self, dialog):
        libs = getattr(dialog, "libs_data", None)
        if not libs:
            messagebox.showinfo("提示", "没有可复制的库列表。")
            return
        text = self._make_libs_text(libs)
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            messagebox.showinfo("提示", "库清单已复制到剪贴板！")
        except Exception as e:
            messagebox.showerror("复制失败", f"复制到剪贴板失败：\n{e}")

    def _build_all_libs_rows(self, scroll_frame, dialog, libs):
        for w in scroll_frame.winfo_children():
            w.destroy()
        dialog.vars_map.clear()
        dialog.row_refs.clear()
        dialog.libs_data = list(libs)

        sorted_libs = [(item, item[0].lower() == "pip") for item in self._libs_sort(libs)]

        try:
            if hasattr(dialog, "count_label"):
                dialog.count_label.config(text=f"共 {len(libs)} 个已安装的库：")
        except Exception:
            pass

        col_min = [150, 90, 0, 110, 90]

        header = tk.Frame(scroll_frame, bg="#e8e8e8")
        header.pack(fill=tk.X, pady=(0, 4))
        header.columnconfigure(0, weight=0, minsize=col_min[0])
        header.columnconfigure(1, weight=0, minsize=col_min[1])
        header.columnconfigure(2, weight=1)
        header.columnconfigure(3, weight=0, minsize=col_min[3])
        header.columnconfigure(4, weight=0, minsize=col_min[4])

        tk.Checkbutton(header, text="库名", font=("微软雅黑", 9), bg="#e8e8e8", anchor="w", state="disabled", disabledforeground="black").grid(row=0, column=0, sticky="w")
        tk.Label(header, text="版本", font=("微软雅黑", 9, "bold"), bg="#e8e8e8", anchor="w").grid(row=0, column=1, sticky="w")
        tk.Label(header, text="说明", font=("微软雅黑", 9, "bold"), bg="#e8e8e8", anchor="w").grid(row=0, column=2, sticky="w")
        tk.Label(header, text="状态", font=("微软雅黑", 9, "bold"), bg="#e8e8e8", anchor="w").grid(row=0, column=3, sticky="w", padx=(24, 0))
        tk.Label(header, text="操作", font=("微软雅黑", 9, "bold"), bg="#e8e8e8", anchor="w").grid(row=0, column=4, sticky="w", padx=(36, 0))

        for (name, ver), is_pip in sorted_libs:
            frame = tk.Frame(scroll_frame)
            frame.pack(fill=tk.X, pady=1)
            frame.columnconfigure(0, weight=0, minsize=col_min[0])
            frame.columnconfigure(1, weight=0, minsize=col_min[1])
            frame.columnconfigure(2, weight=1)
            frame.columnconfigure(3, weight=0, minsize=col_min[3])
            frame.columnconfigure(4, weight=0, minsize=col_min[4])

            var = tk.BooleanVar(value=False)
            if is_pip:
                cb = tk.Checkbutton(frame, text=name, variable=var, font=("微软雅黑", 9), anchor="w", state="disabled")
            else:
                cb = tk.Checkbutton(frame, text=name, variable=var, font=("微软雅黑", 9), anchor="w")
            cb.grid(row=0, column=0, sticky="w")
            dialog.vars_map[name] = (var, cb)

            tk.Label(frame, text=f"{ver}", font=("微软雅黑", 9), anchor="w").grid(row=0, column=1, sticky="w")

            desc = get_lib_description(name)
            tk.Label(frame, text=f"{desc}", font=("微软雅黑", 8), anchor="w").grid(row=0, column=2, sticky="w")

            status_label = tk.Label(frame, text="✅ 已安装", font=("微软雅黑", 8), fg="green")
            status_label.grid(row=0, column=3, sticky="w", padx=(24, 0))

            if is_pip:
                btn = tk.Button(frame, text="卸载", font=("微软雅黑", 8), width=6, state="disabled")
            else:
                btn = tk.Button(frame, text="卸载", font=("微软雅黑", 8), width=6, command=lambda n=name: self._uninstall_one_lib(n, dialog))
            btn.grid(row=0, column=4, sticky="w", padx=(22, 0))

            dialog.row_refs[name] = {"status": status_label, "btn": btn, "var": var, "cb": cb}

    def _all_libs_select(self, dialog, select):
        for name, (var, cb) in dialog.vars_map.items():
            try:
                if cb["state"] == "disabled":
                    continue
            except Exception:
                continue
            var.set(select)

    def _uninstall_one_lib(self, name, dialog):
        if self.busy:
            messagebox.showinfo("提示", "有任务正在进行，请稍候。")
            return
        if not messagebox.askyesno("确认卸载", f"确定要卸载 {name} 吗？"):
            return
        self._do_uninstall([name], dialog)

    def _batch_uninstall(self, dialog):
        if self.busy:
            messagebox.showinfo("提示", "有任务正在进行，请稍候。")
            return
        selected = [n for n, (var, cb) in dialog.vars_map.items() if var.get()]
        if not selected:
            messagebox.showinfo("提示", "请至少勾选一个要卸载的库。")
            return
        if not messagebox.askyesno("确认卸载", f"确定要卸载选中的 {len(selected)} 个库吗？"):
            return
        self._do_uninstall(selected, dialog)

    def _do_uninstall(self, names, dialog):
        self._log("=" * 60)
        self._log(f"🗑️ 开始卸载（共 {len(names)} 个）")
        self._log("-" * 60)

        self._set_busy(True)

        def work():
            for name in names:
                self._log(f"⏳ 正在卸载 {name} ...")

                def callback(line):
                    self._log(line)

                output, code = run_command(
                    PYTHON_CMD + ["-m", "pip", "uninstall", "-y", name],
                    callback=callback,
                )
                if code == 0:
                    self._log(f"✅ {name} 卸载完成")
                    self.root.after(0, lambda n=name: self._mark_row_uninstalled(dialog, n))
                else:
                    self._log(f"❌ {name} 卸载失败（返回码 {code}）")

            self._log("")
            self._log("✅ 卸载操作完成，可点「刷新」查看结果。")
            self._log("=" * 60)
            self._log("")

            def finish():
                self._set_busy(False)
            self.root.after(0, finish)

        threading.Thread(target=work, daemon=True).start()

    def _mark_row_uninstalled(self, dialog, name):
        try:
            refs = dialog.row_refs.get(name)
            if not refs:
                return
            refs["status"].config(text="已卸载", fg="#888888")
            refs["btn"].config(state="disabled")
            refs["var"].set(False)
        except Exception:
            pass

    def _refresh_all_libs(self, dialog):
        _, _, env_ok = self._get_env_status(force=True)
        if not env_ok:
            messagebox.showwarning("环境异常", "⚠️ 当前系统未检测到 Python 环境。")
            return
        libs = get_installed_libs()
        if not libs:
            messagebox.showwarning("提示", "没有检测到已安装的库。")
            return
        self._print_all_libs_to_report(libs)
        self._build_all_libs_rows(dialog.scroll_frame, dialog, libs)


    def _build_export_header(self):
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lines = [
            "# ============================================================",
            "# Python 环境依赖清单",
            "# ------------------------------------------------------------",
            f"# 本文件由「{APP_NAME} {VERSION}」自动生成",
            f"# 作者：{AUTHOR}",
            f"# 生成时间：{now_str}",
            "# ============================================================",
            "#",
            "# 【这个文件是干什么的】",
            "# 它记录了你当前 Python 环境里安装的所有第三方库和版本号。",
            "# 换电脑、重装系统、或在另一台机器上复现同样的环境时，",
            "# 可以用它一键把这一整套库全部装回来。",
            "#",
            "# 【怎么用这个文件恢复环境】",
            "#",
            "# ── 方式一：用本工具恢复（推荐，适合新手）──",
            "#   1. 打开「PY更新安装打包助手」",
            "#   2. 点击左侧「安装区」→「从清单恢复库」",
            "#   3. 选中本文件，确认，等待安装完成即可",
            "#",
            "# ── 方式二：用命令行恢复（不用本工具也行）──",
            "#   1. 打开 CMD（命令提示符，Win+R 输入 cmd 回车）",
            "#   2. 输入下面的命令并回车（把路径换成你本文件的实际路径）：",
            '#        pip install -r "D:\\你的路径\\requirements.txt"',
            "#   3. 等待下载安装完成即可",
            "#",
            "#   ▸ 想装到指定的 Python 环境，可以这样写：",
            '#        C:\\你的Python路径\\python.exe -m pip install -r "D:\\你的路径\\requirements.txt"',
            "#",
            "#   ▸ 国内下载慢？加上 -i 参数换清华源：",
            '#        pip install -r "D:\\你的路径\\requirements.txt" -i https://pypi.tuna.tsinghua.edu.cn/simple',
            "#",
            "# ============================================================",
            "# 以下是依赖清单正文：",
            "# ------------------------------------------------------------",
            "",
        ]
        return "\n".join(lines)

    def _export_requirements(self):
        _, _, env_ok = self._get_env_status()
        if not env_ok:
            messagebox.showwarning("无法导出", "⚠️ 当前系统未检测到 Python 环境。\n\n请点击「安装 Python」进行安装，安装完成后重新打开本工具。")
            return
        libs = get_installed_libs()
        if not libs:
            messagebox.showwarning("提示", "没有检测到已安装的库，无需导出。")
            return
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        default_name = f"requirements_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        file_path = filedialog.asksaveasfilename(initialdir=desktop, initialfile=default_name, defaultextension=".txt", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if not file_path:
            self._log("❌ 取消导出")
            return
        self._log(f"⏳ 正在导出到 {file_path} ...")
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(self._build_export_header())
                f.write("\n")
                for name, ver in sorted(libs):
                    desc = get_lib_description(name)
                    f.write(f"# {desc}\n")
                    f.write(f"{name}=={ver}\n\n")
            self._log(f"✅ 导出成功！文件保存在：{file_path}")
            self._log("")
        except Exception as e:
            self._log(f"❌ 导出失败：{e}")


    def _install_python_guide(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("安装 Python 指引")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.withdraw()

        tk.Label(dialog, text="请前往 Python 官网页面，点击黄色按钮", font=("微软雅黑", 10)).pack(pady=(15, 5))
        tk.Label(dialog, text='"Download Python install manager" 下载安装包并安装。', font=("微软雅黑", 10)).pack(pady=(0, 5))

        url_frame = tk.Frame(dialog)
        url_frame.pack(pady=5)
        python_url = "https://www.python.org/downloads/"
        url_entry = tk.Entry(url_frame, width=45, font=("Consolas", 10))
        url_entry.insert(0, python_url)
        url_entry.config(state="readonly", readonlybackground="#f0f0f0")
        url_entry.pack(side=tk.LEFT, padx=5)

        def open_url():
            webbrowser.open(python_url)

        tk.Button(url_frame, text="打开网址", command=open_url).pack(side=tk.LEFT, padx=5)

        tk.Label(dialog, text="", font=("微软雅黑", 4)).pack()
        tk.Label(dialog, text="⚠️ 注意：目前 Python 官网提供的标准安装包是 .msix 格式。", font=("微软雅黑", 9), fg="#cc6600").pack(pady=(5, 5))
        tk.Label(dialog, text="如果您需要 .exe 格式的安装包，请点击黄色大按钮正下方的那行小字：", font=("微软雅黑", 9)).pack()
        tk.Label(dialog, text='"Or get the standalone installer for Python 3.XX.X"', font=("微软雅黑", 9), fg="blue").pack()
        tk.Label(dialog, text="点击其中的版本号即可下载 .exe 格式的独立安装包。", font=("微软雅黑", 9)).pack(pady=(2, 5))

        tk.Label(dialog, text="", font=("微软雅黑", 2)).pack()
        tk.Label(dialog, text="或者前往以下网址：", font=("微软雅黑", 9)).pack()
        alt_url_frame = tk.Frame(dialog)
        alt_url_frame.pack(pady=3)
        windows_url = "https://www.python.org/downloads/windows/"
        alt_url_entry = tk.Entry(alt_url_frame, width=45, font=("Consolas", 10))
        alt_url_entry.insert(0, windows_url)
        alt_url_entry.config(state="readonly", readonlybackground="#f0f0f0")
        alt_url_entry.pack(side=tk.LEFT, padx=5)

        def open_alt_url():
            webbrowser.open(windows_url)

        tk.Button(alt_url_frame, text="打开备选网址", command=open_alt_url).pack(side=tk.LEFT, padx=5)

        tk.Label(dialog, text="在页面顶部找到最新版本的 Python，点击该版本对应的", font=("微软雅黑", 9)).pack(pady=(5, 0))
        tk.Label(dialog, text='"Windows installer (64-bit)" 下载链接。', font=("微软雅黑", 9)).pack()
        tk.Label(dialog, text="", font=("微软雅黑", 4)).pack()
        tk.Label(dialog, text="⚠️ 安装时请务必勾选 'Add Python to PATH'", font=("微软雅黑", 9), fg="red").pack(pady=(5, 2))
        tk.Label(dialog, text="安装完成后，请重新打开本工具。", font=("微软雅黑", 9)).pack(pady=(0, 10))

        def confirm_and_exit():
            dialog.destroy()
            self.root.quit()
            self.root.destroy()

        tk.Button(dialog, text="确定", command=confirm_and_exit, width=10).pack(pady=(0, 15))

        dialog.update_idletasks()
        need_w = max(580, dialog.winfo_reqwidth())
        need_h = dialog.winfo_reqheight()
        self._center_window(self.root, dialog, need_w, need_h)
        dialog.deiconify()
        dialog.grab_set()

    def _restore_from_file(self):
        _, _, env_ok = self._get_env_status()
        if not env_ok:
            messagebox.showwarning("环境异常", "⚠️ 当前系统未检测到 Python 环境。\n\n请点击「安装 Python」进行安装，安装完成后重新打开本工具。")
            return
        file_path = filedialog.askopenfilename(title="选择 requirements.txt 文件", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if not file_path:
            self._log("❌ 未选择文件")
            return
        self._log("=" * 60)
        self._log(f"📥 开始从清单恢复库：{file_path}")
        self._log("-" * 60)
        if not os.path.exists(file_path):
            self._log("❌ 文件不存在")
            return

        self._set_busy(True)

        def install_thread():
            try:
                self._log("⏳ 正在安装，请稍候...")
                self._log("")

                def callback(line):
                    self._log(line)

                output, code = run_command(
                    PYTHON_CMD + ["-m", "pip", "install", "-r", file_path],
                    callback=callback,
                )

                self._log("")
                if code == 0:
                    self._log("✅ 所有库安装完成！")
                else:
                    self._log(f"⚠️ 安装完成，部分库可能失败（返回码 {code}）")
                self._log("=" * 60)
                self._log("")
            finally:
                def finish():
                    self._set_busy(False)
                    self._refresh_status(force=True)
                self.root.after(0, finish)

        threading.Thread(target=install_thread, daemon=True).start()

    def _install_common_libs(self):
        _, _, env_ok = self._get_env_status()
        if not env_ok:
            messagebox.showwarning("环境异常", "⚠️ 当前系统未检测到 Python 环境。\n\n请点击「安装 Python」进行安装，安装完成后重新打开本工具。")
            return
        installed = [name for name, _ in get_installed_libs()]
        installed_lower = {name.lower() for name in installed}

        dialog = tk.Toplevel(self.root)
        dialog.title("安装常用库")
        self._center_window(self.root, dialog, 500, 450)
        dialog.transient(self.root)
        dialog.grab_set()

        top_frame = tk.Frame(dialog)
        top_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(top_frame, text="请勾选要安装的库：", font=("微软雅黑", 10, "bold")).pack(side=tk.LEFT)
        tk.Button(top_frame, text="全选", command=lambda: self._select_all_common(dialog, True)).pack(side=tk.RIGHT, padx=2)
        tk.Button(top_frame, text="反选", command=lambda: self._select_all_common(dialog, False)).pack(side=tk.RIGHT, padx=2)

        list_frame = tk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        canvas = tk.Canvas(list_frame, highlightthickness=0)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=canvas.yview)
        scroll_frame = tk.Frame(canvas)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas_frame = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        scroll_frame.bind("<Configure>", on_frame_configure)

        def on_mousewheel(event):
            try:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass

        def bind_wheel(e=None):
            canvas.bind_all("<MouseWheel>", on_mousewheel)

        def unbind_wheel(e=None):
            try:
                canvas.unbind_all("<MouseWheel>")
            except Exception:
                pass

        canvas.bind("<Enter>", bind_wheel)
        canvas.bind("<Leave>", unbind_wheel)

        def on_destroy(e):
            if e.widget is dialog:
                unbind_wheel()
        dialog.bind("<Destroy>", on_destroy)

        vars_map = {}
        for name in COMMON_LIB_NAMES:
            desc = get_lib_description(name)
            frame = tk.Frame(scroll_frame)
            frame.pack(fill=tk.X, pady=1)
            is_installed = name.lower() in installed_lower
            var = tk.BooleanVar(value=is_installed)
            if is_installed:
                status = "✅ 已安装"
                fg_color = "green"
                cb = tk.Checkbutton(frame, text=name, variable=var, font=("微软雅黑", 9), state="disabled", anchor="w")
            else:
                status = "⬜ 未安装"
                fg_color = "gray"
                cb = tk.Checkbutton(frame, text=name, variable=var, font=("微软雅黑", 9), anchor="w")
            cb.pack(side=tk.LEFT)
            tk.Label(frame, text=f"  {desc}", font=("微软雅黑", 9), fg="gray").pack(side=tk.LEFT, padx=5)
            tk.Label(frame, text=status, font=("微软雅黑", 8), fg=fg_color, width=10).pack(side=tk.RIGHT)
            vars_map[name] = (var, cb)

        dialog.vars_map = vars_map

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        def do_install():
            selected = [name for name, (var, cb) in vars_map.items() if var.get() and name.lower() not in installed_lower]
            if not selected:
                messagebox.showinfo("提示", "没有选中任何需要安装的库（已安装的库不会被重复安装）。")
                return
            dialog.destroy()
            self._log("=" * 60)
            self._log(f"📥 开始安装选中库（共 {len(selected)} 个）")
            self._log("-" * 60)

            self._set_busy(True)

            def install_thread():
                try:
                    for name in selected:
                        self._log(f"⏳ 正在安装 {name} ...")

                        def callback(line):
                            self._log(line)

                        output, code = run_command(
                            PYTHON_CMD + ["-m", "pip", "install", name],
                            callback=callback,
                        )
                        if code == 0:
                            self._log(f"✅ {name} 安装成功")
                        else:
                            self._log(f"❌ {name} 安装失败")
                    self._log("")
                    self._log("✅ 所有选中的库安装完成！")
                    self._log("=" * 60)
                    self._log("")
                finally:
                    def finish():
                        self._set_busy(False)
                        self._refresh_status(force=True)
                    self.root.after(0, finish)

            threading.Thread(target=install_thread, daemon=True).start()

        tk.Button(btn_frame, text="安装选中", font=("微软雅黑", 10), bg="#4CAF50", fg="white", command=do_install).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text="取消", font=("微软雅黑", 10), command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

        dialog._resize_job = None

        def apply_canvas_width(w):
            dialog._resize_job = None
            try:
                canvas.itemconfig(canvas_frame, width=w)
            except Exception:
                pass

        def on_resize(event):
            w = max(event.width - 10, 1)
            if dialog._resize_job is not None:
                try:
                    dialog.after_cancel(dialog._resize_job)
                except Exception:
                    pass
            dialog._resize_job = dialog.after(40, lambda: apply_canvas_width(w))
        canvas.bind("<Configure>", on_resize)

    def _select_all_common(self, dialog, select):
        for name, (var, cb) in dialog.vars_map.items():
            var.set(select)


    def _print_update_result_to_report(self, pip_current, pip_latest, has_pip_update, outdated_others, all_libs):
        update_items = []
        if has_pip_update:
            update_items.append(("pip", pip_current, pip_latest))
        for lib in outdated_others:
            update_items.append((lib["name"], lib["current"], lib["latest"]))

        updated_names = {name for name, _, _ in update_items}
        latest_items = [(name, ver) for name, ver in all_libs if name not in updated_names]
        latest_items.sort(key=lambda x: x[0].lower())

        self._log("=" * 60)
        self._log("📦 库更新检查结果")
        self._log("-" * 60)

        if update_items:
            self._log(f"可更新（{len(update_items)} 个）：")
            for name, cur, lat in update_items:
                desc = get_lib_description(name)
                self._log(f"  {name}=={cur} → {lat}    # {desc}")
            self._log("")
        else:
            self._log("可更新：无")
            self._log("")

        if latest_items:
            self._log(f"已是最新（{len(latest_items)} 个）：")
            for name, ver in latest_items:
                desc = get_lib_description(name)
                self._log(f"  {name}=={ver}    # {desc}")
        self._log("")
        self._log("=" * 60)
        self._log("")

    def _check_lib_update(self):
        _, _, env_ok = self._get_env_status()
        if not env_ok:
            messagebox.showwarning("环境异常", "⚠️ 当前系统未检测到 Python 环境。\n\n请点击「安装 Python」进行安装，安装完成后重新打开本工具。")
            return

        self._log("=" * 60)
        self._log("📦 正在检查库更新...")
        self._log("-" * 60)

        self._set_busy(True)

        def check_thread():
            try:
                pip_current = self._get_pip_version()

                self._log("⏳ 正在获取库列表...")
                self._log("")

                output_lines = []

                def callback(line):
                    self._log(line)
                    output_lines.append(line)

                output, code = run_command(
                    PYTHON_CMD + ["-m", "pip", "list", "--outdated"],
                    callback=callback,
                    timeout=120,
                )

                if code == -2:
                    self._log("")
                    self._log("⚠️ 检查超时：pip 响应超过 120 秒。")
                    self._log("   可能是网络太慢，或 pip 源无法访问。")
                    self._log("=" * 60)
                    self._log("")

                    def report_timeout():
                        self._set_busy(False)
                        messagebox.showerror(
                            "检查超时",
                            "⚠️ 检查库更新超时（超过 120 秒）。\n\n"
                            "可能原因：\n"
                            "  • 网络连接太慢\n"
                            "  • pip 源无法访问\n\n"
                            "请检查网络后重试。"
                        )
                    self.root.after(0, report_timeout)
                    return

                if code != 0:
                    self._log("")
                    self._log("⚠️ 检查失败：无法获取更新信息。")
                    self._log("   可能原因：网络不通、pip 异常、或 pip 源无法访问。")
                    self._log("=" * 60)
                    self._log("")

                    def report_fail():
                        self._set_busy(False)
                        messagebox.showerror(
                            "检查失败",
                            "⚠️ 无法获取库更新信息。\n\n"
                            "可能原因：\n"
                            "  • 网络连接不通\n"
                            "  • pip 源无法访问\n"
                            "  • pip 本身异常\n\n"
                            "请检查网络后重试。"
                        )
                    self.root.after(0, report_fail)
                    return

                pip_latest = None
                outdated_others = []
                for line in output_lines:
                    if "Package" in line or "Version" in line or "----" in line:
                        continue
                    if not line.strip():
                        continue
                    if "[notice]" in line or "[warning]" in line or "[error]" in line:
                        continue
                    parts = line.split()
                    if len(parts) >= 3:
                        name = parts[0]
                        if name.lower() == "pip":
                            pip_latest = parts[2]
                        else:
                            outdated_others.append({
                                "name": name,
                                "current": parts[1],
                                "latest": parts[2],
                            })

                has_pip_update = (pip_current is not None and pip_latest is not None and pip_current != pip_latest)
                has_other_update = len(outdated_others) > 0
                has_any_update = has_pip_update or has_other_update

                self._log("")
                if has_any_update:
                    total = (1 if has_pip_update else 0) + len(outdated_others)
                    self._log(f"✅ 发现 {total} 个库有新版本可用。")
                else:
                    self._log("✅ 所有库已是最新版本！")
                self._log("=" * 60)
                self._log("")

                all_libs = get_installed_libs()

                self._print_update_result_to_report(
                    pip_current, pip_latest, has_pip_update,
                    outdated_others, all_libs
                )

                def show():
                    self._set_busy(False)
                    self._show_update_dialog(
                        pip_current, pip_latest, has_pip_update,
                        outdated_others, has_any_update, all_libs
                    )
                self.root.after(0, show)
            except Exception as e:
                def report_err():
                    self._set_busy(False)
                    messagebox.showerror("检查异常", f"检查库更新时出错：\n{e}")
                self.root.after(0, report_err)

        threading.Thread(target=check_thread, daemon=True).start()

    def _show_update_dialog(self, pip_current, pip_latest, has_pip_update, outdated_others, has_any_update, all_libs):
        if outdated_others is None:
            outdated_others = []
        if all_libs is None:
            all_libs = []

        dialog = tk.Toplevel(self.root)
        dialog.title("📦 库更新检查")
        self._center_window(self.root, dialog, 680, 480)
        dialog.transient(self.root)
        dialog.grab_set()

        top_frame = tk.Frame(dialog)
        top_frame.pack(fill=tk.X, padx=10, pady=5)
        if has_any_update:
            total_update_count = 0
            if has_pip_update:
                total_update_count += 1
            total_update_count += len(outdated_others)
            tk.Label(top_frame, text=f"以下 {total_update_count} 个库有新版本可用：", font=("微软雅黑", 10, "bold")).pack(side=tk.LEFT)
        else:
            tk.Label(top_frame, text=f"✅ 所有 {len(all_libs)} 个库已是最新版本！", font=("微软雅黑", 10, "bold"), fg="green").pack(side=tk.LEFT)

        list_frame = tk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        canvas = tk.Canvas(list_frame, highlightthickness=0)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=canvas.yview)
        scroll_frame = tk.Frame(canvas)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        canvas_frame = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        scroll_frame.bind("<Configure>", on_frame_configure)

        def on_mousewheel(event):
            try:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass

        def bind_wheel(e=None):
            canvas.bind_all("<MouseWheel>", on_mousewheel)

        def unbind_wheel(e=None):
            try:
                canvas.unbind_all("<MouseWheel>")
            except Exception:
                pass

        canvas.bind("<Enter>", bind_wheel)
        canvas.bind("<Leave>", unbind_wheel)

        def on_destroy(e):
            if e.widget is dialog:
                unbind_wheel()
        dialog.bind("<Destroy>", on_destroy)

        vars_map = {}
        if has_any_update:
            if has_pip_update:
                frame = tk.Frame(scroll_frame)
                frame.pack(fill=tk.X, pady=2)
                tk.Label(frame, text="●", font=("微软雅黑", 12), fg="#cc6600").pack(side=tk.LEFT, padx=(0, 3))
                tk.Label(frame, text="pip", font=("微软雅黑", 9, "bold"), width=14, anchor="w").pack(side=tk.LEFT)
                tk.Label(frame, text=f"{pip_current}", font=("微软雅黑", 9), fg="#666666", width=12).pack(side=tk.LEFT)
                tk.Label(frame, text="→", font=("微软雅黑", 9), fg="#999999").pack(side=tk.LEFT)
                tk.Label(frame, text=f"{pip_latest}", font=("微软雅黑", 9), fg="#cc6600", width=12).pack(side=tk.LEFT)
                desc = get_lib_description("pip")
                if desc:
                    tk.Label(frame, text=f"  {desc}", font=("微软雅黑", 8), fg="#888888").pack(side=tk.LEFT, padx=3)
                tk.Label(frame, text="[需单独更新]", font=("微软雅黑", 8), fg="red").pack(side=tk.LEFT, padx=5)
                sep_line = tk.Frame(scroll_frame, height=2, bg="#cccccc")
                sep_line.pack(fill=tk.X, pady=6)

            for lib in outdated_others:
                frame = tk.Frame(scroll_frame)
                frame.pack(fill=tk.X, pady=2)
                var = tk.BooleanVar(value=True)
                vars_map[lib["name"]] = (var, lib)
                cb = tk.Checkbutton(frame, text=lib["name"], variable=var, font=("微软雅黑", 9), anchor="w", width=14)
                cb.pack(side=tk.LEFT)
                tk.Label(frame, text=f"{lib['current']}", font=("微软雅黑", 9), fg="#666666", width=12).pack(side=tk.LEFT)
                tk.Label(frame, text="→", font=("微软雅黑", 9), fg="#999999").pack(side=tk.LEFT)
                tk.Label(frame, text=f"{lib['latest']}", font=("微软雅黑", 9), fg="green", width=12).pack(side=tk.LEFT)
                desc = get_lib_description(lib["name"])
                if desc:
                    tk.Label(frame, text=f"  {desc}", font=("微软雅黑", 8), fg="#888888").pack(side=tk.LEFT, padx=5)

            total_installed = len(all_libs)
            updated_count = 0
            if has_pip_update:
                updated_count += 1
            updated_count += len(outdated_others)
            other_count = total_installed - updated_count
            if other_count > 0:
                tk.Label(scroll_frame, text="", font=("微软雅黑", 4)).pack()
                tk.Label(scroll_frame, text=f"✅ 其他 {other_count} 个库已是最新版本", font=("微软雅黑", 9), fg="green").pack(anchor="w", padx=5, pady=2)
        else:
            for name, ver in sorted(all_libs):
                frame = tk.Frame(scroll_frame)
                frame.pack(fill=tk.X, pady=1)
                tk.Label(frame, text=f"{name}", font=("微软雅黑", 9), width=20, anchor="w").pack(side=tk.LEFT)
                tk.Label(frame, text=f"{ver}", font=("微软雅黑", 9), fg="#666666", width=12).pack(side=tk.LEFT)
                tk.Label(frame, text="✅", font=("微软雅黑", 9), fg="green").pack(side=tk.LEFT, padx=5)
                desc = get_lib_description(name)
                if desc:
                    tk.Label(frame, text=f"  {desc}", font=("微软雅黑", 8), fg="#888888").pack(side=tk.LEFT, padx=5)

        dialog.vars_map = vars_map

        bottom_frame = tk.Frame(dialog)
        bottom_frame.pack(fill=tk.X, padx=10, pady=10)

        if has_any_update:
            if has_pip_update:
                tk.Label(bottom_frame, text="💡 pip 是 Python 的包管理工具，需单独更新。点击下方【更新pip】按钮。", font=("微软雅黑", 9), fg="#666666").pack(anchor="w", pady=(0, 8))
            btn_frame = tk.Frame(bottom_frame)
            btn_frame.pack(side=tk.RIGHT)
            if has_pip_update:
                tk.Button(btn_frame, text="更新pip", font=("微软雅黑", 10), bg="#cc6600", fg="white", command=lambda: self._update_pip(dialog)).pack(side=tk.RIGHT, padx=5)
            if outdated_others:
                tk.Button(btn_frame, text="更新选中", font=("微软雅黑", 10), bg="#4CAF50", fg="white", command=lambda: self._do_update_selected_from_dialog(dialog)).pack(side=tk.RIGHT, padx=5)
                tk.Button(btn_frame, text="全部更新", font=("微软雅黑", 10), bg="#2196F3", fg="white", command=lambda: self._do_update_all_from_dialog(dialog)).pack(side=tk.RIGHT, padx=5)
            tk.Button(btn_frame, text="取消", font=("微软雅黑", 10), command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        else:
            tk.Button(bottom_frame, text="确定", font=("微软雅黑", 10), width=10, command=dialog.destroy).pack(side=tk.RIGHT)

        dialog._resize_job = None

        def apply_canvas_width(w):
            dialog._resize_job = None
            try:
                canvas.itemconfig(canvas_frame, width=w)
            except Exception:
                pass

        def on_resize(event):
            w = max(event.width - 10, 1)
            if dialog._resize_job is not None:
                try:
                    dialog.after_cancel(dialog._resize_job)
                except Exception:
                    pass
            dialog._resize_job = dialog.after(40, lambda: apply_canvas_width(w))
        canvas.bind("<Configure>", on_resize)

    def _update_pip(self, dialog):
        if dialog:
            dialog.destroy()
        self._log("=" * 60)
        self._log("📦 开始更新 pip")
        self._log("-" * 60)

        self._set_busy(True)

        def callback(line):
            self._log(line)

        def update_thread():
            try:
                output, code = run_command(
                    PYTHON_CMD + ["-m", "pip", "install", "--upgrade", "pip"],
                    callback=callback,
                )
                self._log("")
                if code == 0:
                    self._log("✅ pip 更新成功！")
                else:
                    self._log(f"❌ pip 更新失败（返回码 {code}）")
                self._log("=" * 60)
                self._log("")
            finally:
                def finish():
                    self._set_busy(False)
                    self._refresh_status(force=True)
                    if code == 0:
                        messagebox.showinfo("pip 更新完成", "✅ pip 已更新完成！\n\n请重新点击「检查库更新」查看其他库的更新。")
                    else:
                        messagebox.showerror("pip 更新失败", f"❌ pip 更新失败（返回码 {code}）")
                self.root.after(0, finish)

        threading.Thread(target=update_thread, daemon=True).start()

    def _do_update_selected_from_dialog(self, dialog):
        selected = [lib for name, (var, lib) in dialog.vars_map.items() if var.get()]
        if not selected:
            messagebox.showinfo("提示", "请至少勾选一个要更新的库。")
            return
        dialog.destroy()
        self._log("=" * 60)
        self._log(f"📦 开始更新选中的库（共 {len(selected)} 个）")
        self._log("-" * 60)

        self._set_busy(True)

        def update_thread():
            try:
                for lib in selected:
                    self._log(f"⏳ 正在更新 {lib['name']} ...")

                    def callback(line):
                        self._log(line)

                    output, code = run_command(
                        PYTHON_CMD + ["-m", "pip", "install", "--upgrade", lib["name"]],
                        callback=callback,
                    )
                    if code == 0:
                        self._log(f"✅ {lib['name']} 更新成功: {lib['current']} → {lib['latest']}")
                    else:
                        self._log(f"❌ {lib['name']} 更新失败")
                self._log("")
                self._log("✅ 所有选中的库更新完成！")
                self._log("=" * 60)
                self._log("")
            finally:
                def finish():
                    self._set_busy(False)
                    self._refresh_status(force=True)
                self.root.after(0, finish)

        threading.Thread(target=update_thread, daemon=True).start()

    def _do_update_all_from_dialog(self, dialog):
        all_libs = [lib for _, lib in dialog.vars_map.values()]
        if not all_libs:
            messagebox.showinfo("提示", "没有可更新的库。")
            return
        dialog.destroy()
        self._log("=" * 60)
        self._log(f"📦 开始更新全部库（共 {len(all_libs)} 个）")
        self._log("-" * 60)

        self._set_busy(True)

        def update_thread():
            try:
                for lib in all_libs:
                    self._log(f"⏳ 正在更新 {lib['name']} ...")

                    def callback(line):
                        self._log(line)

                    output, code = run_command(
                        PYTHON_CMD + ["-m", "pip", "install", "--upgrade", lib["name"]],
                        callback=callback,
                    )
                    if code == 0:
                        self._log(f"✅ {lib['name']} 更新成功: {lib['current']} → {lib['latest']}")
                    else:
                        self._log(f"❌ {lib['name']} 更新失败")
                self._log("")
                self._log("✅ 全部库更新完成！")
                self._log("=" * 60)
                self._log("")
            finally:
                def finish():
                    self._set_busy(False)
                    self._refresh_status(force=True)
                self.root.after(0, finish)

        threading.Thread(target=update_thread, daemon=True).start()


    def _select_script(self):
        file_path = filedialog.askopenfilename(title="选择 Python 脚本 (.py)", filetypes=[("Python files", "*.py"), ("All files", "*.*")])
        if file_path:
            self.script_path.set(file_path)
            self._log(f"📄 已选择脚本：{file_path}")

    def _toggle_icon(self):
        if self.opt_icon.get():
            self.icon_entry.config(state="normal")
            self.icon_btn.config(state="normal")
        else:
            self.icon_entry.config(state="disabled")
            self.icon_btn.config(state="disabled")
            self.icon_path.set("")

    def _select_icon(self, var):
        file_path = filedialog.askopenfilename(title="选择图标文件 (.ico)", filetypes=[("Icon files", "*.ico"), ("All files", "*.*")])
        if file_path:
            var.set(file_path)
            self._log(f"🖼️ 已选择图标：{file_path}")

    def _toggle_n_icon(self):
        if self.opt_n_icon.get():
            self.n_icon_entry.config(state="normal")
            self.n_icon_btn.config(state="normal")
        else:
            self.n_icon_entry.config(state="disabled")
            self.n_icon_btn.config(state="disabled")
            self.n_icon_path.set("")

    def _check_nuitka_env(self):
        _, _, env_ok = self._get_env_status()
        if not env_ok:
            messagebox.showwarning("环境异常", "⚠️ 当前系统未检测到 Python 环境。\n\n请点击「安装 Python」进行安装，安装完成后重新打开本工具。")
            return
        self._log("=" * 60)
        self._log("🔍 检测 Nuitka 环境")
        self._log("-" * 60)
        self._log("⏳ 正在检测 Nuitka 环境，请稍候...")

        self._set_busy(True)

        def work():
            try:
                nuitka_ver = None
                timed_out = False
                try:
                    r = subprocess.run(
                        PYTHON_CMD + ["-m", "nuitka", "--version"],
                        capture_output=True, text=True, encoding="utf-8",
                        errors="replace", creationflags=CREATE_NO_WINDOW,
                        env=_get_sub_env(), stdin=subprocess.DEVNULL, timeout=120,
                    )
                    if r.returncode == 0 and r.stdout.strip():
                        nuitka_ver = r.stdout.strip().splitlines()[0]
                except subprocess.TimeoutExpired:
                    timed_out = True
                except Exception:
                    pass

                def done():
                    self._set_busy(False)
                    if timed_out:
                        self._log("⏱️ 检测超时：Nuitka 响应异常，请检查 Python 环境后重试")
                    elif nuitka_ver:
                        self._log(f"✅ Nuitka 已安装，版本：{nuitka_ver}")
                        compiler_ver = self._find_c_compiler()
                        if compiler_ver:
                            self._log(f"✅ 编译组件已安装，版本：{compiler_ver}")
                            self._log("   环境齐全，可以直接打包了")
                        else:
                            self._log("❌ 缺少编译组件（约98MB）")
                            self._log("   💡 Nuitka 打包必须要有这个编译组件，目前电脑上没有检测到。")
                            self._log("   安装步骤：")
                            self._log("   ① 打开 CMD 命令行窗口")
                            self._log("   ② 复制下面这条命令，在命令行窗口ctrl+v粘贴，然后回车：")
                            self._log("      python -m nuitka --version")
                            self._log("   ③ 安装过程中询问 yes 还是 no 的时候，输入 yes 回车")
                            self._log("   ④ 等它下载完，看到 \"Version C compiler\" 就是装好了")
                            self._log("   装好后回来再点一次「检测Nuitka环境」确认")
                    else:
                        self._log("❌ Nuitka 未安装")
                        self._log(" 💡 小提示：Nuitka 打包不是必须的，新手用 PyInstaller 打包就够用了，Nuitka 的库和依赖可以不装。除非您知道这是做什么的。")
                        self._log("   安装步骤（照着做就行）：")
                        self._log("   ① CMD打开一个命令行窗口")
                        self._log("   ② 复制这行命令：")
                        self._log("      pip install nuitka")
                        self._log("   ③ Ctrl + V 粘贴，然后回车")
                        self._log("   ④ 等它显示 Successfully installed nuitka-4.X.X(版本号) 就是装好了")
                        self._log("   💡提示：Nuitka 打包，必须同时具备 Nuitka 软件 和 ziglang 编译工具，二者缺一不可。但是检测 [编译工具 ziglang] 是否安装到位的前提是系统必须安装有[Nuitka 软件],所以请在装好 Nuitka 后,重启软件，再点一次「检测Nuitka环境」按钮，软件就可以检查编译工具 ziglang的安装情况了,如果没装，软件会给您提示，按照提示一步步装好就可以了.")
                    self._log("=" * 60)
                    self._log("")

                self.root.after(0, done)
            except Exception as e:
                def report_err():
                    self._set_busy(False)
                    self._log(f"❌ 检测出错：{e}")
                self.root.after(0, report_err)

        threading.Thread(target=work, daemon=True).start()

    def _clear_nuitka_env(self):
        target = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Nuitka")

        ok, reason = is_safe_nuitka_cache(target)
        if not ok:
            self._log("⛔ 安全检查未通过，已拒绝清除")
            self._log(f"   原因：{reason}")
            messagebox.showwarning(
                "安全检查未通过",
                f"⛔ 出于安全考虑，已拒绝清除操作。\n\n原因：{reason}",
            )
            return

        if not os.path.isdir(target):
            self._log("=" * 60)
            self._log("🗑️ 清除 Nuitka 环境")
            self._log("-" * 60)
            self._log("✅ 未发现 Nuitka 缓存目录，无需清除")
            self._log("=" * 60)
            self._log("")
            return

        total = 0
        try:
            for dirpath, _, filenames in os.walk(target):
                for fn in filenames:
                    try:
                        total += os.path.getsize(os.path.join(dirpath, fn))
                    except Exception:
                        pass
        except Exception:
            pass
        size_mb = total / (1024.0 * 1024.0)

        confirmed = messagebox.askyesno(
            "清除 Nuitka 环境",
            "⚠️ 此操作将永久删除以下目录，删除后不可恢复！\n\n"
            f"{target}\n\n"
            f"占用空间：约 {size_mb:.1f} MB\n\n"
            "该目录里是 Nuitka 自动下载的编译工具（Dependency Walker、\n"
            "C 编译器）和编译缓存。\n\n"
            "删除后：\n"
            "  · 已经装好的 Nuitka 软件本身不受影响\n"
            "  · 下次打包时，Nuitka 会重新下载这些组件\n\n"
            "确认要删除吗？",
            icon="warning",
            default="no",
        )
        if not confirmed:
            self._log("已取消清除操作")
            return

        self._set_busy(True)
        self._log("=" * 60)
        self._log("🗑️ 清除 Nuitka 环境")
        self._log("-" * 60)
        self._log(f"   目标目录：{target}")
        self._log(f"   占用空间：约 {size_mb:.1f} MB")

        def work():
            err = None
            try:
                shutil.rmtree(target)
            except Exception as e:
                err = e

            def done():
                self._set_busy(False)
                if err is None:
                    self._log("✅ 已清除 Nuitka 环境，下载内容已全部删除")
                    self._log("   下次打包时，Nuitka 会重新下载所需组件")
                else:
                    self._log(f"❌ 清除失败：{err}")
                    self._log("   可能有文件正被占用，请关闭其他程序后重试")
                    parent = os.path.dirname(target)
                    if os.path.isdir(parent):
                        self._open_folder(parent)
                self._log("=" * 60)
                self._log("")

            self.root.after(0, done)

        threading.Thread(target=work, daemon=True).start()

    def _find_c_compiler(self):
        for name in ("zig.exe", "gcc.exe", "clang.exe"):
            for d in os.environ.get("PATH", "").split(os.pathsep):
                if not d:
                    continue
                exe = os.path.join(d, name)
                if os.path.isfile(exe):
                    ver = self._read_compiler_version(exe)
                    return ver or name
        cache_root = os.path.join(
            os.environ.get("LOCALAPPDATA", ""),
            "Nuitka", "Nuitka", "Cache", "downloads",
        )
        if os.path.isdir(cache_root):
            for dirpath, _, filenames in os.walk(cache_root):
                for fn in filenames:
                    if fn.lower() in ("zig.exe", "gcc.exe", "clang.exe"):
                        exe = os.path.join(dirpath, fn)
                        ver = self._read_compiler_version(exe)
                        return ver or fn
        return None

    def _read_compiler_version(self, exe):
        try:
            r = subprocess.run(
                [exe, "--version"],
                capture_output=True, text=True,
                encoding="utf-8", errors="replace",
                creationflags=CREATE_NO_WINDOW, timeout=10,
            )
            if r.returncode == 0:
                first = (r.stdout or r.stderr).strip().splitlines()
                if first:
                    return first[0].strip()
        except Exception:
            pass
        return None

    def _ask_cleanup(self, temp_dir, failed=False):
        if not temp_dir or not os.path.isdir(temp_dir):
            return

        ok, reason = is_safe_to_delete(temp_dir)
        if not ok:
            self._log(f"⛔ 安全检查未通过，拒绝清理：{temp_dir}")
            self._log(f"   原因：{reason}")
            return

        try:
            if failed:
                keep = messagebox.askyesno(
                    "打包失败",
                    f"打包失败了。\n\n"
                    f"本工具自己创建的临时目录：\n{temp_dir}\n\n"
                    f"保留它有助于排查问题（里面有编译日志）。\n\n"
                    f"要保留吗？\n"
                    f"选「是」= 保留　选「否」= 移入回收站"
                )
            else:
                clean = messagebox.askyesno(
                    "打包完成",
                    f"打包成功！\n\n"
                    f"本工具自己创建的临时目录：\n{temp_dir}\n\n"
                    f"这些临时文件已经没用了，可以移入回收站释放空间。\n\n"
                    f"要现在清理吗？"
                )
                keep = not clean
        except Exception:
            return

        if keep:
            self._log(f"📁 已保留临时目录：{temp_dir}")
        else:
            ok, msg = safe_delete_dir(temp_dir)
            if ok:
                self._log(f"🗑️ {msg}")
            else:
                self._log(f"⚠️ 清理未执行：{msg}")
                self._open_folder(temp_dir)
                try:
                    messagebox.showinfo("需要手动清理", msg)
                except Exception:
                    pass

    def _prepare_script_for_pack(self, script, temp_dir):
        if getattr(sys, "frozen", False):
            return script, False
        try:
            is_self = os.path.abspath(script) == os.path.abspath(__file__)
        except Exception:
            return script, False
        if not is_self:
            return script, False

        copy_dir = os.path.join(temp_dir, "源文件副本")
        try:
            os.makedirs(copy_dir, exist_ok=True)
            copy_path = os.path.join(copy_dir, os.path.basename(script))
            shutil.copy2(script, copy_path)
        except Exception as e:
            self._log(f"⚠️ 自动复制源文件失败：{e}")
            self._log("   将直接从原文件打包（可能因文件被占用而失败）。")
            return script, False

        self._log("📋 检测到正在打包本工具自身，已自动复制一份到临时目录：")
        self._log(f"   {copy_path}")
        self._log("   将从副本打包，避免运行时文件被占用导致失败。")
        return copy_path, True

    def _start_pack_nuitka(self):
        py_ver, _, _ = self._get_env_status()
        if not py_ver:
            messagebox.showwarning("环境异常", "⚠️ 未检测到 Python 环境，无法进行打包操作。\n\n请点击「安装 Python」进行安装，安装完成后重新打开本工具。")
            self._log("❌ 打包失败：未检测到 Python 环境")
            return

        script = self.script_path.get().strip()
        if not script:
            messagebox.showwarning("提示", "请先选择一个 Python 脚本。")
            return
        if not os.path.exists(script):
            messagebox.showwarning("提示", "脚本文件不存在，请重新选择。")
            return
        if not is_lib_installed("nuitka"):
            messagebox.showwarning("缺少打包库", "⚠️ 未检测到打包库 nuitka\n\n请点击「检测Nuitka环境」按钮，按提示一步步安装 Nuitka，\n装好后重启软件，再重新点击「开始打包」。")
            return

        if not self._find_c_compiler():
            messagebox.showwarning(
                "缺少编译组件",
                "⚠️ 检测到电脑上缺少 Nuitka 打包用的编译组件（约98MB）。\n\n"
                "请点击「检测Nuitka环境」按钮，软件会检查编译组件，\n"
                "并提示您一步步安装，装好后重新点击「开始打包」。",
            )
            return

        if self.opt_n_icon.get():
            icon = self.n_icon_path.get().strip()
            if not icon or not os.path.exists(icon):
                messagebox.showwarning("提示", "已勾选自定义图标，请先选择图标文件。")
                return
            if not icon.lower().endswith(".ico"):
                messagebox.showwarning("提示", "图标文件必须为 .ico 格式。")
                return

        save_path = filedialog.asksaveasfilename(title="保存 EXE 文件到", defaultextension=".exe", filetypes=[("Executable files", "*.exe")], initialfile=os.path.splitext(os.path.basename(script))[0])
        if not save_path:
            self._log("❌ 取消打包")
            return

        output_dir = os.path.dirname(os.path.abspath(save_path))
        output_name = os.path.splitext(os.path.basename(save_path))[0]

        nuitka_temp_dir = os.path.join(output_dir, "Nuitka临时文件")
        nuitka_out_dir = os.path.join(nuitka_temp_dir, "output")
        try:
            os.makedirs(nuitka_out_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("无法创建临时目录", f"无法创建：\n{nuitka_out_dir}\n\n{e}")
            return

        script_to_pack, _used_copy = self._prepare_script_for_pack(script, nuitka_temp_dir)

        cmd = PYTHON_CMD + ["-m", "nuitka"]
        cmd.append("--assume-yes-for-downloads")
        if self.opt_n_onefile.get():
            cmd.append("--onefile")
        if self.opt_n_windowed.get():
            cmd.append("--windows-console-mode=disable")
        cmd.append("--enable-plugin=tk-inter")
        if self.opt_n_icon.get() and self.n_icon_path.get().strip():
            cmd.append(f"--windows-icon-from-ico={self.n_icon_path.get().strip()}")
        cmd.append(f"--output-filename={output_name}.exe")
        cmd.append(f"--output-dir={nuitka_out_dir}")
        cmd.append("--remove-output")
        cmd.append(script_to_pack)

        self._log("=" * 60)
        self._log("🚀 开始打包（Nuitka）")
        self._log("-" * 60)
        self._log(f"脚本: {script}")
        self._log(f"最终输出目录: {output_dir}")
        self._log(f"临时目录（本工具自建）: {nuitka_temp_dir}")
        self._log("")
        self._log("⏳ 正在打包，Nuitka 首次编译较慢，请耐心等待...")
        self._log("")

        self._set_busy(True)

        def pack_thread():
            try:
                def callback(line):
                    self._log(line)
                output, code = run_command(cmd, callback=callback)
                self._log("")

                if code == 0:
                    src_exe = os.path.join(nuitka_out_dir, f"{output_name}.exe")
                    dst_exe = os.path.join(output_dir, f"{output_name}.exe")
                    moved = False
                    try:
                        if os.path.exists(src_exe):
                            if os.path.abspath(src_exe) != os.path.abspath(dst_exe):
                                if os.path.exists(dst_exe):
                                    try:
                                        os.remove(dst_exe)
                                    except Exception:
                                        pass
                                shutil.move(src_exe, dst_exe)
                            moved = True
                    except Exception as e:
                        self._log(f"⚠️ EXE 移动失败：{e}")
                        self._log(f"   EXE 仍在临时目录：{nuitka_out_dir}")

                    if moved and os.path.exists(dst_exe):
                        self._log(f"✅ 打包成功！EXE 位于：{dst_exe}")
                    elif os.path.exists(dst_exe):
                        self._log(f"✅ 打包成功！EXE 位于：{dst_exe}")
                    else:
                        self._log("✅ 打包完成，但未在临时目录中找到 EXE。")
                        self._log(f"   请查看：{nuitka_out_dir}")
                else:
                    self._log(f"❌ 打包失败（返回码 {code}）")
                    self._log(f"   临时目录保留在：{nuitka_temp_dir}")
                self._log("=" * 60)
                self._log("")

                def finish():
                    self._set_busy(False)
                    self._ask_cleanup(nuitka_temp_dir, failed=(code != 0))
                self.root.after(0, finish)
            except Exception as e:
                def report_err():
                    self._set_busy(False)
                    self._log(f"❌ 打包异常：{e}")
                self.root.after(0, report_err)

        threading.Thread(target=pack_thread, daemon=True).start()

    def _start_pack(self):
        py_ver, _, _ = self._get_env_status()
        if not py_ver:
            messagebox.showwarning("环境异常", "⚠️ 未检测到 Python 环境，无法进行打包操作。\n\n请点击「安装 Python」进行安装，安装完成后重新打开本工具。")
            self._log("❌ 打包失败：未检测到 Python 环境")
            return

        script = self.script_path.get().strip()
        if not script:
            messagebox.showwarning("提示", "请先选择一个 Python 脚本。")
            return
        if not os.path.exists(script):
            messagebox.showwarning("提示", "脚本文件不存在，请重新选择。")
            return
        if not is_lib_installed("pyinstaller"):
            messagebox.showwarning("缺少打包库", "⚠️ 未检测到打包库 pyinstaller\n\n请前往左侧「安装区」→「安装常用库」，\n勾选 pyinstaller 进行安装。\n\n安装完成后重新点击「开始打包」。")
            return

        if self.opt_icon.get():
            icon = self.icon_path.get().strip()
            if not icon or not os.path.exists(icon):
                messagebox.showwarning("提示", "已勾选自定义图标，请先选择图标文件。")
                return
            if not icon.lower().endswith(".ico"):
                messagebox.showwarning("提示", "图标文件必须为 .ico 格式。")
                return

        save_path = filedialog.asksaveasfilename(title="保存 EXE 文件到", defaultextension=".exe", filetypes=[("Executable files", "*.exe")], initialfile=os.path.splitext(os.path.basename(script))[0])
        if not save_path:
            self._log("❌ 取消打包")
            return

        output_dir = os.path.dirname(os.path.abspath(save_path))
        output_name = os.path.splitext(os.path.basename(save_path))[0]

        temp_dir = os.path.join(output_dir, "打包临时文件")
        try:
            os.makedirs(temp_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("无法创建临时目录", f"无法创建：\n{temp_dir}\n\n{e}")
            return

        script_to_pack, _used_copy = self._prepare_script_for_pack(script, temp_dir)

        cmd = PYTHON_CMD + ["-m", "PyInstaller"]
        if self.opt_onefile.get():
            cmd.append("--onefile")
        if self.opt_windowed.get():
            cmd.append("--windowed")
        if self.opt_icon.get() and self.icon_path.get().strip():
            cmd.extend(["--icon", self.icon_path.get().strip()])
        cmd.extend(["--name", output_name])
        cmd.extend(["--distpath", output_dir])
        cmd.extend(["--workpath", os.path.join(temp_dir, "build")])
        cmd.extend(["--specpath", temp_dir])
        cmd.append(script_to_pack)

        self._log("=" * 60)
        self._log("🚀 开始打包")
        self._log("-" * 60)
        self._log(f"脚本: {script}")
        self._log(f"输出目录: {output_dir}")
        self._log(f"临时目录（本工具自建）: {temp_dir}")
        self._log("")
        self._log("⏳ 正在打包，请等待...")
        self._log("")

        self._set_busy(True)

        def pack_thread():
            try:
                def callback(line):
                    self._log(line)
                output, code = run_command(cmd, callback=callback)
                self._log("")
                if code == 0:
                    exe_path = os.path.join(output_dir, f"{output_name}.exe")
                    if os.path.exists(exe_path):
                        self._log(f"✅ 打包成功！EXE 位于：{exe_path}")
                    else:
                        self._log("✅ 打包完成，EXE 文件已生成。")
                else:
                    self._log(f"❌ 打包失败（返回码 {code}）")
                self._log("=" * 60)
                self._log("")

                def finish():
                    self._set_busy(False)
                    self._ask_cleanup(temp_dir, failed=(code != 0))
                self.root.after(0, finish)
            except Exception as e:
                def report_err():
                    self._set_busy(False)
                    self._log(f"❌ 打包异常：{e}")
                self.root.after(0, report_err)

        threading.Thread(target=pack_thread, daemon=True).start()

    def _show_about(self):
        msg = (f"关于 {APP_NAME}\n\n一个让新手也能轻松管理 Python 环境的小工具。\n\n作者：{AUTHOR}\n版本：{VERSION}\n基于内部测试版 2.9.3 发布\n制作日期：{self.build_date}\n\n功能：\n  • 检测 Python 环境状态\n  • 一键安装 / 修复常用库\n  • 查看所有库并支持单个 / 批量卸载\n  • 库清单可复制到报告区或剪贴板\n  • 从 requirements.txt 恢复全部依赖\n  • 将 Python 脚本打包成 EXE\n  • 检查 Python 库是否有更新\n\n安全特性：\n  • 打包临时目录只清理工具自建的子文件夹，绝不碰你的其他文件\n  • 清理走回收站，可恢复\n  • 自己打包自己时自动复制，静默处理\n\n基于 Python + Tkinter 开发")
        messagebox.showinfo("关于", msg)



WINDOW_ICON_B64 = (
    "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"
)

TASKBAR_APP_ID = "sunny15007.PYEnvPackAssistant.2.9.3"


def _set_taskbar_identity():
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(TASKBAR_APP_ID)
    except Exception:
        pass


def _apply_window_icon(root):
    try:
        ico_path = os.path.join(tempfile.gettempdir(), "PYEnvPackAssistant.ico")
        with open(ico_path, "wb") as f:
            f.write(base64.b64decode(WINDOW_ICON_B64))
        root.iconbitmap(ico_path)
        root.iconbitmap(default=ico_path)
        _set_icon_by_dpi(root, ico_path)
    except Exception:
        pass


def _set_icon_by_dpi(root, ico_path):
    try:
        import ctypes

        user32 = ctypes.windll.user32
        user32.GetParent.argtypes = [ctypes.c_void_p]
        user32.GetParent.restype = ctypes.c_void_p
        user32.LoadImageW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint,
                                      ctypes.c_int, ctypes.c_int, ctypes.c_uint]
        user32.LoadImageW.restype = ctypes.c_void_p
        user32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint,
                                        ctypes.c_size_t, ctypes.c_ssize_t]
        user32.SendMessageW.restype = ctypes.c_ssize_t

        hwnd = user32.GetParent(root.winfo_id()) or root.winfo_id()

        DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4

        cx_big, cx_small = 0, 0
        old_ctx = 0
        try:
            user32.SetThreadDpiAwarenessContext.argtypes = [ctypes.c_void_p]
            user32.SetThreadDpiAwarenessContext.restype = ctypes.c_void_p
            old_ctx = user32.SetThreadDpiAwarenessContext(
                ctypes.c_void_p(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2))
            cx_big = user32.GetSystemMetrics(11)
            cx_small = user32.GetSystemMetrics(49)
        except Exception:
            cx_big, cx_small = 0, 0
        finally:
            if old_ctx:
                try:
                    user32.SetThreadDpiAwarenessContext(ctypes.c_void_p(old_ctx))
                except Exception:
                    pass

        if cx_big <= 0 or cx_small <= 0:
            try:
                dpi = user32.GetDpiForWindow(hwnd)
            except Exception:
                dpi = 0
            if not dpi:
                dpi = 96
            try:
                cx_big = user32.GetSystemMetricsForDpi(11, dpi)
                cx_small = user32.GetSystemMetricsForDpi(49, dpi)
            except Exception:
                cx_big, cx_small = 32, 16
        if cx_big <= 0:
            cx_big = 32
        if cx_small <= 0:
            cx_small = 16

        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x0010
        WM_SETICON = 0x0080
        ICON_SMALL, ICON_BIG = 0, 1

        cx_small = cx_big

        h_big = user32.LoadImageW(None, ico_path, IMAGE_ICON, cx_big, cx_big, LR_LOADFROMFILE)
        h_small = user32.LoadImageW(None, ico_path, IMAGE_ICON, cx_small, cx_small, LR_LOADFROMFILE)
        if h_big:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, h_big)
        if h_small:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, h_small)
    except Exception:
        pass


if __name__ == "__main__":
    _set_taskbar_identity()
    root = tk.Tk()
    root.withdraw()
    _apply_window_icon(root)
    app = PyAssistantApp(root)
    root.mainloop()