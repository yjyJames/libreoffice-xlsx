import os
import time
import subprocess

LIBRE_OFFICE_HOME = os.environ.get("LIBRE_OFFICE_HOME", r"C:\Program Files\LibreOffice")
PROGRAM_DIR = os.path.join(LIBRE_OFFICE_HOME, "program")

# Run this file with LibreOffice's bundled Python so `uno` imports cleanly.
import uno
from com.sun.star.beans import PropertyValue

SOFFICE_PATH = os.path.join(PROGRAM_DIR, "soffice.exe")
XLSX_PATH = r"E:\pythonProject\excel_skill\demo.xlsx"

HOST = "127.0.0.1"
PORT = 2002

# 给 UNO 单独用一个 LibreOffice 用户配置目录，避免和正常打开的 LibreOffice 冲突
LO_PROFILE_DIR = r"C:\temp\lo-uno-profile"


def make_property(name, value):
    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


def start_libreoffice():
    os.makedirs(LO_PROFILE_DIR, exist_ok=True)

    profile_url = uno.systemPathToFileUrl(os.path.abspath(LO_PROFILE_DIR))

    accept_arg = (
        f"--accept=socket,host={HOST},port={PORT};"
        f"urp;"
        f"StarOffice.ComponentContext"
    )

    args = [
        SOFFICE_PATH,
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--norestore",
        "--nolockcheck",
        accept_arg,
    ]

    print("正在启动 LibreOffice...")
    proc = subprocess.Popen(args)

    return proc


def connect_libreoffice(timeout=30):
    local_ctx = uno.getComponentContext()

    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        local_ctx
    )

    uno_url = (
        f"uno:socket,host={HOST},port={PORT};"
        f"urp;"
        f"StarOffice.ComponentContext"
    )

    start_time = time.time()

    while True:
        try:
            ctx = resolver.resolve(uno_url)
            print("已连接 LibreOffice")
            return ctx
        except Exception:
            if time.time() - start_time > timeout:
                raise RuntimeError("连接 LibreOffice 超时")
            time.sleep(0.5)


def main():
    if not os.path.exists(SOFFICE_PATH):
        raise FileNotFoundError(f"找不到 LibreOffice: {SOFFICE_PATH}")

    if not os.path.exists(XLSX_PATH):
        raise FileNotFoundError(f"找不到 xlsx 文件: {XLSX_PATH}")

    proc = start_libreoffice()

    ctx = connect_libreoffice()

    smgr = ctx.ServiceManager

    desktop = smgr.createInstanceWithContext(
        "com.sun.star.frame.Desktop",
        ctx
    )

    file_url = uno.systemPathToFileUrl(os.path.abspath(XLSX_PATH))
    #
    # props = (
    #     make_property("Hidden", False),
    #     make_property("ReadOnly", False),
    # )
    #
    # print("正在打开 xlsx...")
    #
    # doc = desktop.loadComponentFromURL(
    #     file_url,
    #     "_blank",
    #     0,
    #     props
    # )
    #
    # if doc is None:
    #     raise RuntimeError("打开 xlsx 失败，doc is None")
    #
    # print("xlsx 已打开:", XLSX_PATH)

    # 防止 Python 立刻退出，方便你观察 LibreOffice 是否还会闪退
    input("按回车退出 Python，但不会主动关闭 LibreOffice...")


if __name__ == "__main__":
    main()
