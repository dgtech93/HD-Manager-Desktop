from __future__ import annotations

# Thin wrapper around the existing UI implementation.
# The full view logic currently lives in `app.ui.main_window.MainWindow`.
# Keeping this wrapper allows moving imports to MVC paths without breaking behavior.

from app.ui.main_window import MainWindow as _LegacyMainWindow


class MainWindow(_LegacyMainWindow):
    def __init__(self, controller, parent=None) -> None:
        # MVC wiring: the legacy UI still expects an object it calls `repository`,
        # and optionally a `system` attribute for OS actions.
        self.controller = controller
        self.system = getattr(controller, "system", None)
        super().__init__(controller)

