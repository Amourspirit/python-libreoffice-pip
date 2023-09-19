from __future__ import annotations
import os


def is_flatpak() -> bool:
    """Detect if LibreOffice is running in a Flatpak sandbox."""
    fp_id = os.getenv("FLATPAK_ID", "")
    if fp_id == "org.libreoffice.LibreOffice":
        return True
    return False
