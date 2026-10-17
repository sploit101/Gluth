#!/usr/bin/env python3
"""Run the local FastAPI backend and Vite frontend together."""
from __future__ import annotations

import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FRONTEND = ROOT / "frontend"


def main() -> int:
    npm = shutil.which("npm.cmd" if sys.platform == "win32" else "npm")
    if npm is None:
        print("npm is required. Install Node.js first.", file=sys.stderr)
        return 2
    if not (FRONTEND / "node_modules").exists():
        print("Run `npm install` in frontend/ first.", file=sys.stderr)
        return 2

    processes = [
        subprocess.Popen([sys.executable, "-m", "uvicorn", "serve:app",
                          "--host", "127.0.0.1", "--port", "8800"], cwd=ROOT),
        subprocess.Popen([npm, "run", "dev", "--", "--host", "127.0.0.1",
                          "--port", "5173"], cwd=FRONTEND),
    ]
    print("Frontend: http://127.0.0.1:5173", flush=True)
    print("Backend:  http://127.0.0.1:8800/api/health", flush=True)
    try:
        while all(process.poll() is None for process in processes):
            time.sleep(0.5)
        return next((process.returncode or 1 for process in processes
                     if process.poll() is not None), 1)
    except KeyboardInterrupt:
        return 0
    finally:
        for process in processes:
            if process.poll() is None:
                process.terminate()
        for process in processes:
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()


if __name__ == "__main__":
    raise SystemExit(main())
