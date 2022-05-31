import sys, os
from cx_Freeze import setup, Executable

__version__ = "9.0.0"

include_files = ['data']
packages = ["tkinter", "os", "openpyxl", "selenium"]
# excludes = [""]

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Publicador Yapo",
    author="Jose Miguel",
    description='Sistema de Control',
    version=__version__,
    options={"build_exe": {
        'packages': packages,
        'include_files': include_files,
        # 'excludes': excludes,
        'include_msvcr': False,
    }},
    executables=[Executable("main.py", base=base, icon="data/icon.ico")]
)
