import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os"], "excludes": ["tkinter"]}

# GUI applications require a different base on Windows (the default is for
# a console application).
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name = "nlDIAPE",
    version = "0.4",
    description = "Algoritmo medición del sentimiento, Google Natural Language",
    options = {"build_exe": build_exe_options},
    executables = [Executable("nlDIAPE.py", base=base)]
)