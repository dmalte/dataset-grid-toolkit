# Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

"""QWebChannel bridge — thin adapter between the grid UI and native desktop APIs.

Exposes methods to JavaScript via QWebChannel:
    pick_file(filter)   → native file-open dialog → absolute path
    pick_folder()       → native folder dialog   → absolute path
    read_file(path)     → file read               → content string
    write_file(path, content) → file write         → {success, error}
    get_tools()         → JSON summary of tools.yaml
    resolve_and_run_tool(tool_id, ctx) → resolve + run non-interactive tool
    start_interactive_tool(tool_id, ctx) → start interactive QProcess tool
    send_tool_input(text) → write to stdin of running interactive tool
    cancel_tool()       → kill running interactive tool
    get_setting / set_setting → persisted tool settings

No business logic lives here — the CLI remains the single source of truth.
Allowed executables are derived from tools.yaml declarations (no hardcoded whitelist).
"""

import json
import os
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import QObject, QProcess, QProcessEnvironment, QSettings, Signal, Slot
from PySide6.QtWidgets import QFileDialog

from tools import load_tools_yaml, tools_summary, resolve_tool, get_allowed_executables

_APP_DIR = Path(__file__).resolve().parent


class Bridge(QObject):
    """QObject exposed to the web page via QWebChannel."""

    # Signals for interactive tool output (connected by JS via QWebChannel)
    tool_output = Signal(str)    # stdout/stderr chunk
    tool_finished = Signal(int)  # exit code

    def __init__(self, project_root: Path, parent=None):
        super().__init__(parent)
        self._project_root = project_root
        self._settings = QSettings("DataVisualizationGrid", "Desktop")
        self._last_dir = str(self._settings.value("last_dir", str(project_root)))
        self._tools_yaml = _APP_DIR / "tools.yaml"
        self._tools_cfg = load_tools_yaml(self._tools_yaml) if self._tools_yaml.is_file() else {}
        self._interactive_process: QProcess | None = None
        self._interactive_resolved: dict | None = None
        self._page = None  # Set by main.py after page creation

    def set_page(self, page):
        """Store a reference to the QWebEnginePage for runJavaScript calls."""
        self._page = page

    def _remember_dir(self, path: str):
        self._last_dir = path
        self._settings.setValue("last_dir", path)

    def _resolve_executable(self, name: str) -> str:
        """Resolve an executable name to a full path using the venv Scripts dir."""
        scripts_dir = Path(sys.executable).parent
        # On Windows, entry-point scripts have .exe suffix
        for suffix in ("", ".exe"):
            candidate = scripts_dir / (name + suffix)
            if candidate.is_file():
                return str(candidate)
        return name  # fallback to bare name (rely on PATH)

    def _validate_and_resolve_args(self, args: list[str]) -> tuple[list[str] | None, str]:
        """Validate executable against tools.yaml declarations and resolve to full path.

        Returns (resolved_args, error). On success error is empty.
        """
        executable = args[0] if args else ""
        base_exe = os.path.basename(executable).lower()
        if base_exe.endswith(".exe"):
            base_exe = base_exe[:-4]
        allowed = get_allowed_executables(self._tools_cfg)
        if base_exe not in allowed:
            return None, f"Executable not declared in tools.yaml: {executable}"
        resolved = list(args)
        resolved[0] = self._resolve_executable(base_exe)
        return resolved, ""

    # ------------------------------------------------------------------
    # File dialogs
    # ------------------------------------------------------------------

    @Slot(str, result=str)
    def pick_file(self, filter_str: str = "") -> str:
        """Open a native file dialog and return the selected absolute path.

        *filter_str* uses Qt filter syntax, e.g. ``"JSON (*.json)"``.
        Returns an empty string when the user cancels.
        """
        path, _ = QFileDialog.getOpenFileName(
            None, "Open File", self._last_dir, filter_str or "All Files (*)"
        )
        if path:
            self._remember_dir(str(Path(path).parent))
        return path or ""

    @Slot(str, str, result=str)
    def pick_save_file(self, suggested_name: str = "", filter_str: str = "") -> str:
        """Open a native save-file dialog and return the selected absolute path.

        *suggested_name* pre-fills the filename field.
        *filter_str* uses Qt filter syntax, e.g. ``"JSON (*.json)"``.
        Returns an empty string when the user cancels.
        """
        start_dir = self._last_dir
        if suggested_name:
            start_dir = str(Path(start_dir) / suggested_name)
        path, _ = QFileDialog.getSaveFileName(
            None, "Save File", start_dir, filter_str or "All Files (*)"
        )
        if path:
            self._remember_dir(str(Path(path).parent))
        return path or ""

    @Slot(result=str)
    def pick_folder(self) -> str:
        """Open a native folder dialog and return the selected absolute path."""
        path = QFileDialog.getExistingDirectory(
            None, "Select Folder", self._last_dir
        )
        if path:
            self._remember_dir(path)
        return path or ""

    # ------------------------------------------------------------------
    # Tools
    # ------------------------------------------------------------------

    @Slot(result=str)
    def get_tools(self) -> str:
        """Return a JSON summary of all tools defined in tools.yaml."""
        try:
            return json.dumps(tools_summary(self._tools_cfg))
        except Exception as exc:
            return json.dumps({"error": str(exc)})

    @Slot(str, str, result=str)
    def resolve_and_run_tool(self, tool_id: str, context_json: str) -> str:
        """Resolve tool templates and run the CLI command.

        *context_json* is a JSON object with keys:
            control_values  — {control_id: value} from the UI
            dataset_path    — current dataset file path (or null)
            source_info     — tracked-source metadata dict (or null)

        Returns a JSON object: {stdout, stderr, returncode, outputs, track_source, result}.
        """
        try:
            ctx = json.loads(context_json)
        except (json.JSONDecodeError, TypeError) as exc:
            return json.dumps({"stdout": "", "stderr": str(exc), "returncode": -1})

        control_values = ctx.get("control_values") or {}
        dataset_path = ctx.get("dataset_path")
        source_info = ctx.get("source_info")

        # Collect current settings values
        settings_values = {}
        for key in (self._tools_cfg.get("settings") or {}):
            val = self._settings.value(f"tool_setting_{key}", "")
            if val:
                settings_values[key] = str(val)

        try:
            resolved = resolve_tool(
                self._tools_cfg,
                tool_id,
                control_values=control_values,
                settings_values=settings_values,
                project_root=self._project_root,
                dataset_path=dataset_path,
                source_info=source_info,
            )
        except Exception as exc:
            return json.dumps({"stdout": "", "stderr": f"Tool resolution failed: {exc}", "returncode": -1})

        args = resolved["args"]
        resolved_args, err = self._validate_and_resolve_args(args)
        if err:
            return json.dumps({"stdout": "", "stderr": err, "returncode": -1})

        try:
            result = subprocess.run(
                resolved_args,
                capture_output=True,
                text=True,
                cwd=str(self._project_root),
                timeout=120,
            )
            return json.dumps({
                "stdout": result.stdout,
                "stderr": result.stderr,
                "returncode": result.returncode,
                "outputs": resolved["outputs"],
                "track_source": resolved["track_source"],
                "result": resolved["result"],
            })
        except subprocess.TimeoutExpired:
            return json.dumps(
                {"stdout": "", "stderr": "Command timed out after 120 seconds", "returncode": -1}
            )
        except Exception as exc:
            return json.dumps(
                {"stdout": "", "stderr": str(exc), "returncode": -1}
            )

    # ------------------------------------------------------------------
    # Interactive tools (QProcess-based stdin/stdout piping)
    # ------------------------------------------------------------------

    @Slot(str, str, result=str)
    def start_interactive_tool(self, tool_id: str, context_json: str) -> str:
        """Start an interactive tool as a QProcess with piped stdin/stdout.

        Returns JSON: {ok, error, outputs, track_source, result}.
        The tool_output signal emits stdout/stderr chunks as they arrive.
        The tool_finished signal emits the exit code when the process ends.
        """
        if self._interactive_process is not None:
            return json.dumps({"ok": False, "error": "Another interactive tool is already running"})

        try:
            ctx = json.loads(context_json)
        except (json.JSONDecodeError, TypeError) as exc:
            return json.dumps({"ok": False, "error": str(exc)})

        control_values = ctx.get("control_values") or {}
        dataset_path = ctx.get("dataset_path")
        source_info = ctx.get("source_info")

        settings_values = {}
        for key in (self._tools_cfg.get("settings") or {}):
            val = self._settings.value(f"tool_setting_{key}", "")
            if val:
                settings_values[key] = str(val)

        try:
            resolved = resolve_tool(
                self._tools_cfg,
                tool_id,
                control_values=control_values,
                settings_values=settings_values,
                project_root=self._project_root,
                dataset_path=dataset_path,
                source_info=source_info,
            )
        except Exception as exc:
            return json.dumps({"ok": False, "error": f"Tool resolution failed: {exc}"})

        args = resolved["args"]
        resolved_args, err = self._validate_and_resolve_args(args)
        if err:
            return json.dumps({"ok": False, "error": err})

        self._interactive_resolved = resolved

        proc = QProcess(self)
        proc.setWorkingDirectory(str(self._project_root))
        proc.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)

        # Force unbuffered output so readyReadStandardOutput fires promptly
        env = QProcessEnvironment.systemEnvironment()
        env.insert("PYTHONUNBUFFERED", "1")
        proc.setProcessEnvironment(env)

        proc.readyReadStandardOutput.connect(self._on_interactive_stdout)
        proc.finished.connect(self._on_interactive_finished)

        program = resolved_args[0]
        proc_args = resolved_args[1:]
        proc.start(program, proc_args)

        if not proc.waitForStarted(5000):
            self._interactive_process = None
            self._interactive_resolved = None
            return json.dumps({"ok": False, "error": f"Failed to start: {proc.errorString()}"})

        self._interactive_process = proc
        return json.dumps({
            "ok": True,
            "error": "",
            "outputs": resolved["outputs"],
            "track_source": resolved["track_source"],
            "result": resolved["result"],
        })

    @Slot(str)
    def send_tool_input(self, text: str) -> None:
        """Write text to the stdin of the running interactive tool."""
        if self._interactive_process is not None and self._interactive_process.state() == QProcess.ProcessState.Running:
            self._interactive_process.write((text + "\n").encode())

    @Slot()
    def cancel_tool(self) -> None:
        """Kill the running interactive tool."""
        if self._interactive_process is not None:
            self._interactive_process.kill()

    def _push_to_js(self, js_code: str):
        """Run JavaScript on the page (reliable alternative to QWebChannel signals)."""
        if self._page:
            self._page.runJavaScript(js_code)

    def _on_interactive_stdout(self):
        if self._interactive_process is None:
            return
        try:
            raw = self._interactive_process.readAllStandardOutput()
            data = bytes(raw).decode(errors="replace")
            if data:
                escaped = json.dumps(data)
                self._push_to_js(f"window._onToolOutput && window._onToolOutput({escaped})")
        except Exception as exc:
            escaped = json.dumps(f"[bridge error reading stdout: {exc}]\n")
            self._push_to_js(f"window._onToolOutput && window._onToolOutput({escaped})")

    def _on_interactive_finished(self, exit_code: int, _exit_status):
        # Read any remaining output
        if self._interactive_process is not None:
            try:
                raw = self._interactive_process.readAllStandardOutput()
                remaining = bytes(raw).decode(errors="replace")
                if remaining:
                    escaped = json.dumps(remaining)
                    self._push_to_js(f"window._onToolOutput && window._onToolOutput({escaped})")
            except Exception:
                pass
        self._interactive_process = None
        self._interactive_resolved = None
        self._push_to_js(f"window._onToolFinished && window._onToolFinished({exit_code})")

    # ------------------------------------------------------------------
    # Settings (shared tool preferences)
    # ------------------------------------------------------------------

    @Slot(str, result=str)
    def get_setting(self, key: str) -> str:
        """Return a persisted tool setting value, or empty string."""
        return str(self._settings.value(f"tool_setting_{key}", ""))

    @Slot(str, str)
    def set_setting(self, key: str, value: str) -> None:
        """Persist a tool setting value."""
        self._settings.setValue(f"tool_setting_{key}", value)

    # ------------------------------------------------------------------
    # File I/O
    # ------------------------------------------------------------------

    @Slot(str, result=str)
    def read_file(self, path: str) -> str:
        """Read and return the text content of *path*.

        Returns a JSON object ``{content, error}`` — *content* is the file
        text on success, *error* is non-empty on failure.
        """
        try:
            resolved = Path(path).resolve()
            if not resolved.is_file():
                return json.dumps({"content": "", "error": f"Not a file: {path}"})
            content = resolved.read_text(encoding="utf-8")
            return json.dumps({"content": content, "error": ""})
        except Exception as exc:
            return json.dumps({"content": "", "error": str(exc)})

    @Slot(str, str, result=str)
    def write_file(self, path: str, content: str) -> str:
        """Write *content* to *path*, creating parent directories as needed.

        Returns a JSON object ``{success: bool, error: str}``.
        """
        try:
            resolved = Path(path).resolve()
            resolved.parent.mkdir(parents=True, exist_ok=True)
            resolved.write_text(content, encoding="utf-8")
            size = resolved.stat().st_size
            print(f"[BRIDGE] write_file OK  path={resolved}  size={size}B")
            return json.dumps({"success": True, "error": ""})
        except Exception as exc:
            print(f"[BRIDGE] write_file ERR path={path}  error={exc}")
            return json.dumps({"success": False, "error": str(exc)})
