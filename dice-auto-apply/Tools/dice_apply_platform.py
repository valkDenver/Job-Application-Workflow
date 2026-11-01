"""Everything that differs between macOS and Windows, in one place (design section 15).

Paths of the runtime folder and the venv, finding a suitable Python, keeping the machine awake during a round,
starting a round as a detached process, and checking whether a process id is still alive.
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path

IS_WIN = sys.platform.startswith("win")
IS_MAC = sys.platform == "darwin"
HERE = Path(__file__).resolve().parent          # dice-auto-apply/Tools
FOLDER = HERE.parent                            # dice-auto-apply
WORK = Path.home() / ".job-workflow"
VENV = WORK / "venv"
SUPPORTED_PY = ((3, 12), (3, 13))               # Playwright's greenlet wheels lag new Python releases; the venv uses one of these


def runtime_dir(cfg=None):
    p = os.environ.get("DICE_APPLY_RUNTIME") or (cfg or {}).get("runtime_dir")  # the env var is for tests
    return Path(p).expanduser() if p else WORK / "dice-apply"


def venv_python():
    return VENV / ("Scripts/python.exe" if IS_WIN else "bin/python")


def in_venv():
    try:
        return Path(sys.prefix).resolve() == VENV.resolve()
    except OSError:
        return False


def find_supported_python():
    """A Python 3.12/3.13 executable for creating the venv, or None."""
    cands = []
    if IS_WIN:
        py = shutil.which("py")
        for major, minor in reversed(SUPPORTED_PY):
            if py:
                cands.append([py, f"-{major}.{minor}"])
        local = Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Python"
        for major, minor in reversed(SUPPORTED_PY):
            cands.append([str(local / f"Python{major}{minor}" / "python.exe")])
    else:
        for major, minor in reversed(SUPPORTED_PY):
            cands.append([f"python{major}.{minor}"])
            cands.append([f"/Library/Frameworks/Python.framework/Versions/{major}.{minor}/bin/python3"])
            cands.append([f"/opt/homebrew/bin/python{major}.{minor}"])
            cands.append([f"/usr/local/bin/python{major}.{minor}"])
    for c in cands:
        exe = c[0]
        if not (os.path.isabs(exe) and os.path.exists(exe)) and not shutil.which(exe):
            continue
        try:
            r = subprocess.run(c + ["-c", "import sys;print(sys.version_info[:2])"], capture_output=True, text=True, timeout=20)
        except (OSError, subprocess.SubprocessError):
            continue
        if r.returncode == 0 and any(str(v) in r.stdout for v in SUPPORTED_PY):
            return c
    return None


def pid_alive(pid):
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    if IS_WIN:
        import ctypes
        PROCESS_QUERY_LIMITED_INFORMATION, STILL_ACTIVE = 0x1000, 259
        h = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not h:
            return False
        try:
            code = ctypes.c_ulong()
            ok = ctypes.windll.kernel32.GetExitCodeProcess(h, ctypes.byref(code))
            return bool(ok) and code.value == STILL_ACTIVE
        finally:
            ctypes.windll.kernel32.CloseHandle(h)
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


class KeepAwake:
    """Stops the computer from sleeping while a round runs. The screen may still lock; that is fine."""

    def __init__(self):
        self.proc = None

    def __enter__(self):
        try:
            if IS_WIN:
                import ctypes
                ES_CONTINUOUS, ES_SYSTEM_REQUIRED = 0x80000000, 0x00000001
                ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
            elif IS_MAC and shutil.which("caffeinate"):
                self.proc = subprocess.Popen(["caffeinate", "-i", "-w", str(os.getpid())])
        except Exception:
            pass  # keeping awake is a convenience; a round still works without it
        return self

    def __exit__(self, *exc):
        try:
            if IS_WIN:
                import ctypes
                ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
            elif self.proc:
                self.proc.terminate()
        except Exception:
            pass
        return False


def start_detached(cmd, log_path):
    """Start cmd as its own process that keeps running when the terminal, chat or assistant closes."""
    log = open(log_path, "a", encoding="utf-8")
    kw = {"stdout": log, "stderr": subprocess.STDOUT, "stdin": subprocess.DEVNULL, "cwd": str(FOLDER)}
    if IS_WIN:
        DETACHED_PROCESS, CREATE_NEW_PROCESS_GROUP, CREATE_NO_WINDOW = 0x00000008, 0x00000200, 0x08000000
        kw["creationflags"] = DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP | CREATE_NO_WINDOW
        kw["close_fds"] = True
    else:
        kw["start_new_session"] = True
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
    return subprocess.Popen(cmd, env=env, **kw)


def chrome_installed():
    if IS_WIN:
        roots = [os.environ.get(k, "") for k in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA")]
        return any(os.path.exists(os.path.join(r, "Google", "Chrome", "Application", "chrome.exe")) for r in roots if r)
    if IS_MAC:
        return any(os.path.exists(p) for p in ("/Applications/Google Chrome.app", os.path.expanduser("~/Applications/Google Chrome.app")))
    return bool(shutil.which("google-chrome") or shutil.which("google-chrome-stable"))
