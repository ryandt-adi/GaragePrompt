#!/usr/bin/env python3
"""
theme_config.py
===============
Auto-generated PowerPoint theme colors extracted from template.

Source Template: template.pptx
Theme Name: ADI-Brand-PowerPoint-Template-Standard
Extracted: 2026-01-16 21:36:20

DO NOT EDIT MANUALLY - Regenerate using:
    python theme_extractor.py template.pptx
"""

from typing import Dict, Tuple

# =============================================================================
# RAW THEME COLORS (as extracted from template)
# =============================================================================

THEME_COLORS_RAW: Dict[str, str] = {
    "dk1": "000000",
    "lt1": "FFFFFF",
    "dk2": "0067B9",
    "lt2": "9EA1AE",
    "accent1": "00325C",
    "accent2": "1B9CD0",
    "accent3": "8637BA",
    "accent4": "179963",
    "accent5": "FED141",
    "accent6": "C81A28",
    "hlink": "0067B9",
    "folHlink": "3C4157",
}

# =============================================================================
# SEMANTIC COLOR MAPPING
# =============================================================================

THEME_COLORS: Dict[str, str] = {
    "primary_dark": "000000",
    "primary_light": "FFFFFF",
    "secondary_dark": "0067B9",
    "secondary_light": "9EA1AE",
    "accent1": "00325C",
    "accent2": "1B9CD0",
    "accent3": "8637BA",
    "accent4": "179963",
    "accent5": "FED141",
    "accent6": "C81A28",
    "hyperlink": "0067B9",
    "followed_hyperlink": "3C4157",
}

# =============================================================================
# ACCENT COLOR LIST (for charts and data visualization)
# =============================================================================

ACCENT_COLORS: Tuple[str, ...] = (
    "#00325C",
    "#1B9CD0",
    "#8637BA",
    "#179963",
    "#FED141",
    "#C81A28",
)

# =============================================================================
# FONT SCHEME
# =============================================================================

THEME_FONTS: Dict[str, str] = {
    "majorFont": "Barlow Medium",
    "minorFont": "Barlow",
}

# =============================================================================
# CONVENIENCE CONSTANTS
# =============================================================================

# Primary brand color (accent1)
PRIMARY_COLOR = "#00325C"

# Background colors
DARK_BACKGROUND = "#000000"
LIGHT_BACKGROUND = "#FFFFFF"

# Text colors
TEXT_ON_LIGHT = "#000000"
TEXT_ON_DARK = "#FFFFFF"

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def get_accent_color(index: int) -> str:
    """Get accent color by index (0-5), with wraparound."""
    return ACCENT_COLORS[index % len(ACCENT_COLORS)]


def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    """Convert hex color to RGB tuple."""
    h = hex_color.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """Convert RGB values to hex color string."""
    return f"{r:02X}{g:02X}{b:02X}"
