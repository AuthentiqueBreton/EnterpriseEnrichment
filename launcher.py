import os
import sys
import socket
import threading
import webbrowser
from pathlib import Path


def resource_path(relative_path: str) -> str:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return str(base_path / relative_path)


def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return int(s.getsockname()[1])


def main() -> None:
    from streamlit.web import cli as stcli

    app_path = resource_path("app.py")
    port = int(os.environ.get("STREAMLIT_PORT", find_free_port()))
    url = f"http://127.0.0.1:{port}"

    def open_browser() -> None:
        webbrowser.open_new_tab(url)

    threading.Timer(1.5, open_browser).start()

    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--server.address=127.0.0.1",
        f"--server.port={port}",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false",
        "--server.fileWatcherType=none",
    ]

    sys.exit(stcli.main())


if __name__ == "__main__":
    main()