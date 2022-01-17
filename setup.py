import sys
from cx_Freeze import setup, Executable

# Tools> run setup.py > build > Ok
# Dependencies are automatically detected, but it might need fine tuning.
# "packages": ["os"] is used as example only
# include_folder = ['gui/']

build_exe_options = {"packages": ["os"],
                     "excludes": ["tkinter"],
                     "include_files": ['gui/', 'icon.ico']}


# base="Win32GUI" should be used only for Windows GUI app
base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
    Executable(
        "main.py",
        copyright="Copyright (C) 2021 Sun Fashion",
        base=base,
        icon="icon.ico",
        shortcut_name="Online Helper",
        shortcut_dir="Online Helper",
        target_name="Online Helper.exe"
    ),
]

setup(
    name="Online Helper",
    version="0.1",
    description="All scripts to help in online business",
    options={"build_exe": build_exe_options},
    executables=executables
)
