import sys
import os
import threading
import socket
import time
import signal
import logging
import pyautogui
import qrcode_terminal
from datetime import datetime
from flask import Flask, render_template, request
from flask_socketio import SocketIO
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from InquirerPy import inquirer

# ---------------- Helper for CxFreeze paths ---------------- #
if getattr(sys, "frozen", False):
    # Running from cx_Freeze exe
    base_path = os.path.dirname(sys.executable)
else:
    # Running normally (dev mode)
    base_path = os.path.abspath(".")

app = Flask(
    __name__,
    template_folder=os.path.join(base_path, "templates"),
    static_folder=os.path.join(base_path, "static")
)
socketio = SocketIO(app, cors_allowed_origins="*")

logs = []

# ---------------- Flask Routes ---------------- #
@app.route("/")
def index():
    return render_template("remote.html")

# ---------------- Logging ---------------- #
console = Console()

def log_action(action, level="INFO"):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    logs.append(f"[{level}] [{timestamp}] {action}")
    color = "cyan" if level == "INFO" else "red" if level == "ERROR" else "yellow"
    console.print(f"[bold {color}][{level}][/bold {color}] [{timestamp}] {action}")

# ---------------- SocketIO Events ---------------- #
@socketio.on("next")
def next_slide():
    pyautogui.press("right")
    log_action("Next Slide pressed")

@socketio.on("prev")
def prev_slide():
    pyautogui.press("left")
    log_action("Previous Slide pressed")

@socketio.on("start")
def start_presentation():
    pyautogui.press("f5")
    log_action("Presentation started")

@socketio.on("exit")
def exit_presentation():
    pyautogui.press("esc")
    log_action("Presentation exited")

@socketio.on("blank")
def blank_screen():
    pyautogui.press("w")
    log_action("Blank screen activated")

@socketio.on("black")
def black_screen():
    pyautogui.press("b")
    log_action("Black screen activated")

@socketio.on("stop_server")
def stop_server():
    socketio.emit("server_stopping", {"msg": "Server is shutting down!"})
    console.print("⚠️ Stop server event received. Shutting down...")
    os.kill(os.getpid(), signal.SIGINT)

# ---------------- Helper Functions ---------------- #
def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip

def run_server():
    try:
        log_action("Starting Flask server...")
        logging.getLogger("werkzeug").disabled = True
        socketio.run(app, host="0.0.0.0", port=5000, debug=False, use_reloader=False)
    except Exception as e:
        log_action(f"Server error: {e}", level="ERROR")

def generate_qr(ip, port=5000):
    url = f"http://{ip}:{port}"
    console.print(Panel(f"Access your remote at: [bold green]{url}[/bold green]", title="QR Code URL"))
    qrcode_terminal.draw(url, version=2)

# ---------------- Terminal UI ---------------- #
def show_logs():
    if not logs:
        console.print("[yellow]No logs to display.[/yellow]")
        return
    table = Table(title="Server Logs", show_lines=True)
    table.add_column("Level", style="cyan", no_wrap=True)
    table.add_column("Timestamp", style="green")
    table.add_column("Action", style="magenta")
    for log in logs:
        parts = log.split("] ")
        level = parts[0].strip("[")
        timestamp = parts[1].strip("[]")
        action = "] ".join(parts[2:])
        table.add_row(level, timestamp, action)
    console.print(table)

def show_about_dev():
    console.print(
        Panel.fit(
            """[bold cyan]👨‍💻 Developer Info[/bold cyan]

[green]Name:[/green]  Anshul Wycliffe
[green]Role:[/green]  Python & PySide6 Developer
[green]GitHub:[/green]  https://github.com/anshulwycliffe
[green]Project:[/green]  PowerPoint Remote Control

This tool allows you to remotely control your PowerPoint
presentations using any device connected to the same network.
            """,
            title="ℹ️ About Developer",
            border_style="cyan"
        )
    )

def show_usage():
    console.print(
        Panel.fit(
            """[bold yellow]📖 How to Use[/bold yellow]

1. Make sure your computer (running this script) and
   your mobile device are on the [green]same Wi-Fi / LAN network[/green].

2. Run this program and select [cyan]"Start Server"[/cyan].
   A QR code and URL will be generated.

3. Open the URL on your mobile browser
   (or just scan the QR code).

4. Use the on-screen buttons to:
   -  Start Presentation
   -  Next Slide
   -  Previous Slide
   -  Exit Presentation
   -  Blank / Black Screen

Enjoy controlling your slides remotely!
            """,
            title="🛠 Usage Guide",
            border_style="yellow"
        )
    )

def main_menu():
    server_thread = None

    while True:
        console.print(Panel("[bold magenta]PowerPoint Remote Control[/bold magenta]", expand=False))
        choice = inquirer.select(
            message="Select an option:",
            choices=[
                "Start Server",
                "Show Logs",
                "About Dev",
                "Usage",
                "Exit"
            ],
            cycle=True,
            pointer="->"
        ).execute()

        choice_lower = choice.lower()

        if choice_lower == "start server":
            if server_thread and server_thread.is_alive():
                console.print("[yellow]Server is already running.[/yellow]")
            else:
                server_thread = threading.Thread(target=run_server, daemon=True)
                server_thread.start()
                time.sleep(1)  # Give server a moment to start
                ip = get_local_ip()
                generate_qr(ip)
                log_action("Server started")

        elif choice_lower == "show logs":
            show_logs()

        elif choice_lower == "about dev":
            show_about_dev()

        elif choice_lower == "usage":
            show_usage()

        elif choice_lower == "exit":
            console.print("[bold red]Exiting...[/bold red]")
            break

# ---------------- Entry Point ---------------- #
if __name__ == "__main__":
    main_menu()
