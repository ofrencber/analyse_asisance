import os
import socket
import subprocess
import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent


def resolve_path(path: str) -> str:
    return str((APP_DIR / path).resolve())


def _find_available_port(preferred: int, start: int = 8501, stop: int = 8510) -> int:
    candidates = [preferred] + [port for port in range(start, stop + 1) if port != preferred]
    for port in candidates:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
            except OSError:
                continue
            return port
    raise RuntimeError(f"No available port found in {start}-{stop}.")


if __name__ == "__main__":
    preferred_port = int(os.environ.get("MCDM_PORT", "8501"))
    port = _find_available_port(preferred_port)
    print(f"MCDM Karar Destek Sistemi starting on http://localhost:{port}")
    streamlit_args = [
        "streamlit",
        "run",
        resolve_path("mcdm_app.py"),
        "--server.address",
        "127.0.0.1",
        "--server.port",
        str(port),
        "--global.developmentMode=false",
    ]
    try:
        import streamlit.web.cli as stcli
    except ModuleNotFoundError:
        try:
            completed = subprocess.run(["streamlit", "--version"], check=True, capture_output=True, text=True)
        except Exception as exc:
            raise SystemExit(
                "Streamlit bulunamadı. `pip3 install streamlit` kurun veya Streamlit yüklü Python ortamında çalıştırın."
            ) from exc
        print(completed.stdout.strip() or completed.stderr.strip())
        os.execvp("streamlit", streamlit_args)
    else:
        sys.argv = streamlit_args
        sys.exit(stcli.main())
