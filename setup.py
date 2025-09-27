import sys
from cx_Freeze import setup, Executable

# Dependencies that might not be auto-detected
build_exe_options = {
    "packages": ["os", "socket", "engineio", "pyautogui", "flask", "flask_socketio", "qrcode_terminal", "rich", "InquirerPy"],
    "excludes": ["tkinter"],  # if not used
    "include_files": ["templates/", "static/"],  # include your folders
}


setup(
    name="PowerPointRemote",
    version="1.0",
    description="Remote control for PowerPoint",
    options={"build_exe": build_exe_options},
    executables=[Executable("app.py", base=None)]
)
