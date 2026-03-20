from __future__ import annotations

# Wrapper around existing implementation in `app.ui.settings_window`.

from app.ui.settings_window import SettingsWindow as _LegacySettingsWindow


class SettingsWindow(_LegacySettingsWindow):
    pass

