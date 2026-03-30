# Copyright (c) 2026 Malte Doerper. MIT License. See LICENSE file.

"""PySide6 desktop shell for the Data Visualization Grid.

Embeds the existing grid HTML/JS/CSS in a QWebEngineView and exposes a
QWebChannel bridge for native file dialogs and CLI execution.

Usage:
    python app/main.py
"""

import sys
from pathlib import Path

from PySide6.QtCore import QUrl, Qt, QFile, QIODevice
from PySide6.QtGui import QDesktopServices, QShortcut, QKeySequence
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtWebEngineCore import (
    QWebEnginePage,
    QWebEngineProfile,
    QWebEngineScript,
    QWebEngineSettings,
)
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import QApplication, QMainWindow

from bridge import Bridge

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

_APP_DIR = Path(__file__).resolve().parent
_PROJECT_ROOT = _APP_DIR.parent
_GRID_SRC = _PROJECT_ROOT / "grid" / "src"
_INDEX_HTML = _GRID_SRC / "index.html"
_USER_DATA = _PROJECT_ROOT / ".desktop-data"


# ---------------------------------------------------------------------------
# Custom web page with popup support
# ---------------------------------------------------------------------------

class GridWebPage(QWebEnginePage):
    """QWebEnginePage subclass that opens window.open() popups in native windows."""

    def __init__(self, profile, bridge=None, parent=None):
        super().__init__(profile, parent)
        self._bridge = bridge
        self._popup_windows = []

    def createWindow(self, _window_type):
        """Spawn a native PopupWindow for every window.open() call."""
        popup = PopupWindow(self.profile(), self._bridge)
        self._popup_windows.append(popup)
        popup.show()
        return popup.page()

    def acceptNavigationRequest(self, url, nav_type, is_main_frame):
        """Redirect external URLs to the system browser."""
        if is_main_frame and url.scheme() in ("http", "https"):
            QDesktopServices.openUrl(url)
            return False
        return super().acceptNavigationRequest(url, nav_type, is_main_frame)


class PopupWindow(QMainWindow):
    """Native window for JS popups (Graph View, Compare View, etc.)."""

    def __init__(self, profile, bridge=None):
        super().__init__()
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.resize(1200, 800)

        self._page = GridWebPage(profile, bridge, self)
        self._view = QWebEngineView(self)
        self._view.setPage(self._page)
        self.setCentralWidget(self._view)

        settings = self._page.settings()
        settings.setAttribute(
            QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True
        )
        settings.setAttribute(
            QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True
        )

        # Expose the bridge to popup pages (needed for Compare View)
        if bridge:
            channel = QWebChannel(self._page)
            channel.registerObject("bridge", bridge)
            self._page.setWebChannel(channel)

        self._page.titleChanged.connect(self.setWindowTitle)

    def page(self):
        return self._page


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Data Visualization Grid")
        self.resize(1400, 900)

        # --- Web profile (named → persistent localStorage on disk) ---
        _USER_DATA.mkdir(exist_ok=True)
        profile = QWebEngineProfile("grid-desktop", self)
        profile.setHttpCacheType(QWebEngineProfile.HttpCacheType.NoCache)
        profile.setPersistentStoragePath(str(_USER_DATA))

        # --- QWebChannel bridge (created early for popup support) ---
        self._bridge = Bridge(_PROJECT_ROOT, parent=self)

        # --- Web page & view ---
        self._page = GridWebPage(profile, self._bridge, self)
        self._view = QWebEngineView(self)
        self._view.setPage(self._page)
        self.setCentralWidget(self._view)

        # Enable local content access (needed for file:// loading)
        settings = self._page.settings()
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalStorageEnabled, True)

        # --- QWebChannel setup on main page ---
        self._channel = QWebChannel(self._page)
        self._channel.registerObject("bridge", self._bridge)
        self._page.setWebChannel(self._channel)

        # Give the bridge a page reference for runJavaScript push calls
        self._bridge.set_page(self._page)

        # --- Inject qwebchannel.js before page loads ---
        self._inject_webchannel_script(profile)

        # --- Load the grid UI ---
        url = QUrl.fromLocalFile(str(_INDEX_HTML))
        self._view.load(url)

        # --- Update window title when JS notifies us of file changes ---
        self._page.titleChanged.connect(self._on_title_changed)

        # --- Dev shortcuts ---
        QShortcut(QKeySequence(Qt.Key.Key_F12), self, self._toggle_dev_tools)
        QShortcut(QKeySequence(Qt.Key.Key_F5), self, self._reload_page)

        self._dev_view = None

    def _inject_webchannel_script(self, profile: QWebEngineProfile):
        """Inject qwebchannel.js from Qt resources as a user script."""
        # Importing PySide6.QtWebChannel (done above) registers the
        # :/qtwebchannel/qwebchannel.js resource file.
        res = QFile(":/qtwebchannel/qwebchannel.js")
        if not res.open(QIODevice.ReadOnly | QIODevice.Text):
            raise FileNotFoundError(
                "Could not read :/qtwebchannel/qwebchannel.js from Qt resources. "
                "Make sure PySide6-WebEngine is installed."
            )
        source = bytes(res.readAll()).decode("utf-8")
        res.close()

        script = QWebEngineScript()
        script.setName("qwebchannel")
        script.setSourceCode(source)
        script.setInjectionPoint(QWebEngineScript.InjectionPoint.DocumentCreation)
        script.setWorldId(QWebEngineScript.ScriptWorldId.MainWorld)
        script.setRunsOnSubFrames(False)
        profile.scripts().insert(script)

    def _on_title_changed(self, title):
        if title:
            self.setWindowTitle(f"{title} — Data Visualization Grid")
        else:
            self.setWindowTitle("Data Visualization Grid")

    def _reload_page(self):
        self._page.triggerAction(QWebEnginePage.WebAction.Reload)

    def _toggle_dev_tools(self):
        if self._dev_view is None:
            self._dev_view = QWebEngineView()
            self._dev_view.setWindowTitle("DevTools — Data Visualization Grid")
            self._dev_view.resize(1000, 600)
            self._page.setDevToolsPage(self._dev_view.page())
        if self._dev_view.isVisible():
            self._dev_view.hide()
        else:
            self._dev_view.show()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
