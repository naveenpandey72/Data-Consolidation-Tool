"""
Template configuration and schema definition for GRI 303 Water Consolidation.
Decouples template rules from code so future years or different GRI templates
can be configured easily.
"""

import re

# Standard sheet aliases mapping to canonical sheet keys
SHEET_ALIASES = {
    "303-01_for nhai": "303_01_nhai",
    "303_01_for nhai": "303_01_nhai",
    "303-01 for nhai": "303_01_nhai",
    "303_01 for nhai": "303_01_nhai",
    "303-01_nhai": "303_01_nhai",
    
    "303_01_for contractors": "303_01_contractors",
    "303-01_for contractors": "303_01_contractors",
    "303-01 for contractors": "303_01_contractors",
    "303_01 for contractors": "303_01_contractors",
    "303_01_contractors": "303_01_contractors",
    
    "303-03 _ for nhai": "303_03_nhai",
    "303-03_for nhai": "303_03_nhai",
    "303_03_for nhai": "303_03_nhai",
    "303-03 for nhai": "303_03_nhai",
    "303_03 for nhai": "303_03_nhai",
    "303-03_nhai": "303_03_nhai",
    
    "303_03_for contractors": "303_03_contractors",
    "303-03_for contractors": "303_03_contractors",
    "303-03 for contractors": "303_03_contractors",
    "303_03 for contractors": "303_03_contractors",
    "303_03_contractors": "303_03_contractors",
    
    "303-04_ for nhai": "303_04_nhai",
    "303-04_for nhai": "303_04_nhai",
    "303_04_for nhai": "303_04_nhai",
    "303-04 for nhai": "303_04_nhai",
    "303_04 for nhai": "303_04_nhai",
    "303-04_nhai": "303_04_nhai",
    
    "303_04_for contractors": "303_04_contractors",
    "303-04_for contractors": "303_04_contractors",
    "303-04 for contractors": "303_04_contractors",
    "303_04 for contractors": "303_04_contractors",
    "303_04_contractors": "303_04_contractors",
    
    "303-05 for nhai": "303_05_nhai",
    "303-05_for nhai": "303_05_nhai",
    "303_05_for nhai": "303_05_nhai",
    "303_05 for nhai": "303_05_nhai",
    "303-05_nhai": "303_05_nhai",
    
    "303_05_for contractors": "303_05_contractors",
    "303-05_for contractors": "303_05_contractors",
    "303-05 for contractors": "303_05_contractors",
    "303_05 for contractors": "303_05_contractors",
    "303_05_contractors": "303_05_contractors",
}


def normalize_sheet_name(name: str) -> str:
    """Normalizes a sheet name to a canonical form like '303_01_nhai' or '303_03_contractors'."""
    s = name.strip().lower()
    s = re.sub(r"[\s_\-]+", "_", s)
    # Extract core pattern: 303_01, 303_03, 303_04, 303_05 and nhai/contractor
    disc = ""
    for d in ["303_01", "303_1", "303_03", "303_3", "303_04", "303_4", "303_05", "303_5"]:
        if d in s:
            disc = d.replace("303_1", "303_01").replace("303_3", "303_03").replace("303_4", "303_04").replace("303_5", "303_05")
            break
    
    scope = "contractors" if "contractor" in s else ("nhai" if "nhai" in s else "")
    if disc and scope:
        return f"{disc}_{scope}"
    return s


def clean_region_name(filename_or_str: str) -> str:
    """
    Infers region name from filename, e.g.:
    303_Water_Agra.xlsx -> Agra
    PIU-Bareilly.xlsx -> Bareilly
    Jhansi_Return_v2.xlsx -> Jhansi
    """
    base = filename_or_str.rsplit(".", 1)[0]
    # Remove common prefixes and suffixes
    base = re.sub(r"^(303_water_|water_|gri_303_|piu[\s_\-]+|ro[\s_\-]+)", "", base, flags=re.IGNORECASE)
    base = re.sub(r"([\s_\-]+(?:return|final|v\d+|consolidated|data|\(\d+\)))+$", "", base, flags=re.IGNORECASE)
    base = re.sub(r"[_\-]+", " ", base).strip()
    return base.title() if base else "Unknown Region"
