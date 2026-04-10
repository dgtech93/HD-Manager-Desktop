"""Versione applicazione (allineare a installer.iss, RELEASE_NOTES e scripts/build_release.ps1)."""

from datetime import date

__version__ = "1.0.6"
# Data dell’ultimo rilascio pubblicato (ISO YYYY-MM-DD); aggiornare a ogni release.
__release_date__ = "2026-03-27"

PRODUCT_AUTHOR = "Diego Giotta"
PRODUCT_KIND = "Progetto Indipendente"
PRODUCT_CONTACT_EMAIL = "dgtech93@gmail.com"


def release_date_display_it() -> str:
    """Data ultimo rilascio in formato leggibile (italiano)."""
    try:
        d = date.fromisoformat(__release_date__)
    except ValueError:
        return __release_date__
    months = (
        "gennaio",
        "febbraio",
        "marzo",
        "aprile",
        "maggio",
        "giugno",
        "luglio",
        "agosto",
        "settembre",
        "ottobre",
        "novembre",
        "dicembre",
    )
    return f"{d.day} {months[d.month - 1]} {d.year}"
