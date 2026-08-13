from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


windows_path = ROOT / "Start on Windows.bat"
windows_bytes = windows_path.read_bytes()
windows = windows_bytes.decode("ascii")
windows_lower = windows.lower()

require(b"\r\n" in windows_bytes, "The Windows launcher must use CRLF line endings")
require(b"\n" not in windows_bytes.replace(b"\r\n", b""), "The Windows launcher has mixed line endings")
require("scripts\\pip" not in windows_lower, "The Windows launcher must not invoke pip.exe directly")
require(windows_lower.count("-m pip install") >= 2, "Both Windows dependency paths must use python -m pip")
require('set "venv_py=%~dp0venv\\scripts\\python.exe"' in windows_lower, "The virtual environment Python path is missing")
require('-r "%~dp0requirements.txt"' in windows_lower, "The requirements path must work from directories with spaces")
require('type nul > "%setup_marker%"' in windows_lower, "The Windows setup marker is missing")
require(
    windows_lower.index('type nul > "%setup_marker%"') > windows_lower.index("-m pip install"),
    "The Windows setup marker must be written after dependencies install",
)

compatibility = (ROOT / "Start on Windows with API Key.bat").read_text(encoding="ascii").lower()
require('call "%~dp0start on windows.bat"' in compatibility, "The compatibility launcher must run the fixed launcher")

mac = (ROOT / "Start on Mac.command").read_text(encoding="utf-8").lower()
require("venv/bin/pip install" not in mac, "The Mac launcher must not invoke pip directly")
require(mac.count("python3 -m pip install") >= 3, "Mac install and retry instructions must use python -m pip")

readme = (ROOT / "README.md").read_text(encoding="utf-8").lower()
require("python -m pip install -r requirements.txt" in readme, "README setup must use python -m pip")

print("PASS: both launchers invoke pip through the virtual environment Python")
print("PASS: failed first runs remain retryable because the setup marker is written last")
print("PASS: quoted Windows paths support folders whose names contain spaces")
