"""
Percorsi Qt/DLL per eseguibili PyInstaller (Windows).

Deve essere importato prima di qualsiasi modulo PyQt6, altrimenti Windows può
caricare DLL di sistema errate e generare crash nativi (es. 0xc0000409 in Qt6Core.dll).
"""

from __future__ import annotations

import os
import sys


def apply_pyinstaller_qt_paths() -> None:
    if not getattr(sys, "frozen", False):
        return
    base = getattr(sys, "_MEIPASS", None)
    if not base or not os.path.isdir(base):
        return
    try:
        bin_dir = os.path.join(base, "PyQt6", "Qt6", "bin")
        if os.path.isdir(bin_dir) and hasattr(os, "add_dll_directory"):
            os.add_dll_directory(bin_dir)
        plugins = os.path.join(base, "PyQt6", "Qt6", "plugins")
        if os.path.isdir(plugins):
            os.environ.setdefault("QT_PLUGIN_PATH", plugins)
        platforms = os.path.join(plugins, "platforms")
        if os.path.isdir(platforms):
            os.environ.setdefault("QT_QPA_PLATFORM_PLUGIN_PATH", platforms)
    except OSError:
        pass
