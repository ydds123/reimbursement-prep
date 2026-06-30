"""
OCR 公共模块 — ocr_scans.py / ocr_payments.py 共享的 Tesseract 配置和工具函数。

Tesseract 安装路径优先级：
  1. TESSERACT_CMD 环境变量
  2. 系统 PATH 中的 tesseract（which tesseract）
  3. Windows 默认路径
  4. macOS/Linux 默认路径
"""
import os
import shutil
from pathlib import Path


def find_tesseract() -> str:
    """解析 Tesseract 可执行文件路径。"""
    # 1. 显式环境变量
    env_path = os.environ.get("TESSERACT_CMD")
    if env_path and Path(env_path).exists():
        return env_path

    # 2. 系统 PATH
    which = shutil.which("tesseract")
    if which:
        return which

    # 3. Windows 默认安装路径
    for candidate in [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]:
        if Path(candidate).exists():
            return candidate

    # 4. macOS/Linux 默认路径
    if Path("/usr/local/bin/tesseract").exists():
        return "/usr/local/bin/tesseract"
    if Path("/usr/bin/tesseract").exists():
        return "/usr/bin/tesseract"

    raise FileNotFoundError(
        "找不到 Tesseract 可执行文件。请安装 Tesseract OCR 并确保它在 PATH 中，"
        "或设置环境变量 TESSERACT_CMD 指向 tesseract 可执行文件。"
        "Windows: https://github.com/UB-Mannheim/tesseract/wiki"
    )


# ── 金额容差（元）──
MATCH_TOLERANCE = 2.0


def configure_tesseract():
    """定位并注册 Tesseract 路径到 pytesseract。"""
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = find_tesseract()
