#!/usr/bin/env python3
"""
adi_template_config.py
======================
Centralized configuration for ADI PowerPoint template compliance.
All formatting rules, positions, colors, and constants in one place.

FIXED:
- ADIColorPalette frozen dataclass method issue
- Added proper classmethod decorators
- Fixed type hints

Version: 1.0.1
"""

from dataclasses import dataclass, field
from typing import Dict, Tuple, List, ClassVar
from enum import Enum


# =============================================================================
# THEME COLOR INTEGRATION
# =============================================================================

# Try to load extracted theme colors from template
_EXTRACTED_THEME_AVAILABLE = False
_EXTRACTED_COLORS = {}

try:
    from theme_config import (
        THEME_COLORS_RAW,
        THEME_COLORS as EXTRACTED_THEME_COLORS,
        ACCENT_COLORS as EXTRACTED_ACCENT_COLORS,
    )
    _EXTRACTED_THEME_AVAILABLE = True
    _EXTRACTED_COLORS = THEME_COLORS_RAW
except ImportError:
    pass  # Will use hardcoded defaults


def _get_theme_color(name: str, default: str) -> str:
    """Get color from extracted theme or return default."""
    if _EXTRACTED_THEME_AVAILABLE and name in _EXTRACTED_COLORS:
        return f"#{_EXTRACTED_COLORS[name]}"
    return default





# =============================================================================
# ENUMERATIONS
# =============================================================================

class Confidentiality(Enum):
    """Footer confidentiality levels per ADI guidelines."""
    PUBLIC = "public"
    CONFIDENTIAL = "confidential"
    INTERNAL_ONLY = "internal_only"


class SlideType(Enum):
    """All supported slide types."""
    COVER = "cover"
    SECTION = "section"
    CONTENT = "content"
    TWO_COLUMN = "two_column"
    TABLE = "table"
    CHART = "chart"
    IMAGE = "image"
    CLOSING = "closing"
    BLANK = "blank"


class TableStyle(Enum):
    """Table styles per ADI template."""
    DEFAULT_18PT = "default_18"
    RECOMMENDED_14PT = "recommended_14"
    CENTERED_14PT = "centered_14"
    TITLE_CENTER_14PT = "title_center"


class SectionSlideType(Enum):
    """Section slide variants."""
    SECTION_TITLE = "section"
    KEY_MESSAGE = "key_message"


class ChartColorScheme(Enum):
    """Chart color schemes."""
    LIGHT_BACKGROUND = "light"
    DARK_BACKGROUND = "dark"


# =============================================================================
# COLOR PALETTE (Official ADI Brand Colors)
# FIXED: Removed frozen=True, added proper methods
# =============================================================================

@dataclass
class ADIColorPalette:
    """
    Official ADI Brand Color Palette.
    
    Colors are loaded from extracted theme if available,
    otherwise falls back to hardcoded defaults.
    """
    
    # Primary Brand Blue (from theme accent1 or default)
    PRIMARY_BLUE: str = _get_theme_color('accent1', "#0067B9").lstrip('#')
    
    # Dark Blues (from theme dk1/dk2 or defaults)
    DARK_BLUE: str = _get_theme_color('dk2', "#002855").lstrip('#')
    NAVY: str = "#001A3D"
    DEEP_NAVY: str = "#00132D"
    
    # Light Colors (from theme lt1/lt2 or defaults)
    WHITE: str = _get_theme_color('lt1', "#FFFFFF").lstrip('#')
    OFF_WHITE: str = "#F5F5F5"
    LIGHT_GRAY: str = "#E5E5E5"
    
    # Text Colors
    TEXT_DARK: str = _get_theme_color('dk1', "#333333").lstrip('#')
    TEXT_LIGHT: str = _get_theme_color('lt1', "#FFFFFF").lstrip('#')
    TEXT_GRAY: str = "#666666"
    TEXT_MUTED: str = "#999999"
    
    # Table Colors
    TABLE_HEADER: str = "#B8D4E8"
    TABLE_ALT_ROW: str = "#E8F1F8"
    TABLE_BORDER: str = "#D0D0D0"
    
    # Chart Colors - Light Background (from theme accents or defaults)
    CHART_LIGHT_1: str = _get_theme_color('accent1', "#0067B9").lstrip('#')
    CHART_LIGHT_2: str = _get_theme_color('accent2', "#4D9AD4").lstrip('#')
    CHART_LIGHT_3: str = _get_theme_color('accent3', "#99C7E8").lstrip('#')
    CHART_LIGHT_4: str = _get_theme_color('accent4', "#FF6B35").lstrip('#')
    CHART_LIGHT_5: str = _get_theme_color('accent5', "#2E8B57").lstrip('#')
    CHART_LIGHT_6: str = _get_theme_color('accent6', "#6B5B95").lstrip('#')
    
    # Chart Colors - Dark Background
    CHART_DARK_1: str = "#4D9AD4"
    CHART_DARK_2: str = "#99C7E8"
    CHART_DARK_3: str = "#CCE3F4"
    CHART_DARK_4: str = "#FF9966"
    CHART_DARK_5: str = "#66CDAA"
    CHART_DARK_6: str = "#B8A9C9"
    
    def get_chart_colors(self, scheme: ChartColorScheme) -> Tuple[str, ...]:
        """Get chart color palette for specified scheme."""
        if scheme == ChartColorScheme.LIGHT_BACKGROUND:
            return (
                self.CHART_LIGHT_1, self.CHART_LIGHT_2, self.CHART_LIGHT_3,
                self.CHART_LIGHT_4, self.CHART_LIGHT_5, self.CHART_LIGHT_6
            )
        else:
            return (
                self.CHART_DARK_1, self.CHART_DARK_2, self.CHART_DARK_3,
                self.CHART_DARK_4, self.CHART_DARK_5, self.CHART_DARK_6
            )


# Create singleton instance
COLORS = ADIColorPalette()


# =============================================================================
# SLIDE DIMENSIONS
# =============================================================================

@dataclass
class SlideDimensions:
    """Exact slide dimensions per ADI template."""
    WIDTH_CM: float = 33.87
    HEIGHT_CM: float = 19.05
    WIDTH_INCHES: float = 13.333
    HEIGHT_INCHES: float = 7.5
    
    @property
    def WIDTH_EMU(self) -> int:
        return int(self.WIDTH_INCHES * 914400)
    
    @property
    def HEIGHT_EMU(self) -> int:
        return int(self.HEIGHT_INCHES * 914400)


DIMENSIONS = SlideDimensions()


# =============================================================================
# TYPOGRAPHY SPECIFICATIONS
# =============================================================================

@dataclass
class TypographyConfig:
    """Typography specifications per ADI template."""
    TITLE_FONT: str = "Barlow Medium"
    BODY_FONT: str = "Barlow"
    TITLE_MIN_PT: int = 36
    TITLE_MAX_LINES: int = 3
    
    LEVEL_1_SIZE: int = 18
    LEVEL_1_BULLET: bool = False
    LEVEL_1_INDENT: float = 0
    LEVEL_1_SPACE_AFTER: int = 6
    
    LEVEL_2_SIZE: int = 16
    LEVEL_2_BULLET: bool = True
    LEVEL_2_INDENT: float = 0.75
    LEVEL_2_SPACE_AFTER: int = 6
    
    LEVEL_3_SIZE: int = 14
    LEVEL_3_BULLET: bool = True
    LEVEL_3_INDENT: float = 1.5
    LEVEL_3_SPACE_AFTER: int = 6
    
    BULLET_CHAR: str = "•"
    FOOTER_SIZE: int = 10
    URL_SIZE: int = 14
    SUBTITLE_SIZE: int = 20


TYPOGRAPHY = TypographyConfig()


# =============================================================================
# LOCKED CONTAINER POSITIONS
# =============================================================================

@dataclass
class ContainerPositions:
    """Fixed text container positions per ADI Slide Master."""
    
    COVER_TITLE: Dict[str, float] = field(default_factory=lambda: {
        "left": 2.5, "top": 6.0, "width": 12.0, "height": 5.0
    })
    COVER_SUBTITLE: Dict[str, float] = field(default_factory=lambda: {
        "left": 2.5, "top": 11.5, "width": 12.0, "height": 1.5
    })
    COVER_URL: Dict[str, float] = field(default_factory=lambda: {
        "left": 2.5, "top": 17.0, "width": 10.0, "height": 1.0
    })
    SECTION_TITLE: Dict[str, float] = field(default_factory=lambda: {
        "left": 2.5, "top": 7.5, "width": 12.0, "height": 4.0
    })
    CONTENT_TITLE: Dict[str, float] = field(default_factory=lambda: {
        "left": 2.0, "top": 1.2, "width": 29.0, "height": 2.0
    })
    CONTENT_BODY: Dict[str, float] = field(default_factory=lambda: {
        "left": 2.0, "top": 3.8, "width": 29.0, "height": 12.5
    })
    TWO_COL_LEFT: Dict[str, float] = field(default_factory=lambda: {
        "left": 2.0, "top": 3.8, "width": 14.0, "height": 12.5
    })
    TWO_COL_RIGHT: Dict[str, float] = field(default_factory=lambda: {
        "left": 17.0, "top": 3.8, "width": 14.0, "height": 12.5
    })
    CLOSING_TAGLINE: Dict[str, float] = field(default_factory=lambda: {
        "left": 2.5, "top": 10.0, "width": 28.0, "height": 3.0
    })
    CLOSING_URL: Dict[str, float] = field(default_factory=lambda: {
        "left": 2.5, "top": 13.5, "width": 10.0, "height": 1.0
    })
    LOGO_COVER: Dict[str, float] = field(default_factory=lambda: {
        "left": 29.0, "top": 1.0, "width": 4.0
    })
    LOGO_CONTENT: Dict[str, float] = field(default_factory=lambda: {
        "left": 29.5, "top": 0.8, "width": 3.5
    })
    LOGO_CLOSING: Dict[str, float] = field(default_factory=lambda: {
        "left": 2.5, "top": 2.0, "width": 5.5
    })
    FOOTER: Dict[str, float] = field(default_factory=lambda: {
        "left": 1.5, "top": 18.0, "width": 28.0, "height": 0.7
    })
    SLIDE_NUMBER: Dict[str, float] = field(default_factory=lambda: {
        "left": 31.5, "top": 18.0, "width": 1.5, "height": 0.7
    })
    
    def get(self, key: str) -> Dict[str, float]:
        """Get container by key name."""
        key_map = {
            "cover_title": self.COVER_TITLE,
            "cover_subtitle": self.COVER_SUBTITLE,
            "cover_url": self.COVER_URL,
            "section_title": self.SECTION_TITLE,
            "content_title": self.CONTENT_TITLE,
            "content_body": self.CONTENT_BODY,
            "two_col_left": self.TWO_COL_LEFT,
            "two_col_right": self.TWO_COL_RIGHT,
            "closing_tagline": self.CLOSING_TAGLINE,
            "closing_url": self.CLOSING_URL,
            "logo_cover": self.LOGO_COVER,
            "logo_content": self.LOGO_CONTENT,
            "logo_closing": self.LOGO_CLOSING,
            "footer": self.FOOTER,
            "slide_num": self.SLIDE_NUMBER,
        }
        return key_map.get(key, {})


CONTAINERS = ContainerPositions()


# =============================================================================
# AMP TRIANGLE CONFIGURATION
# =============================================================================

@dataclass
class AMPTriangleConfig:
    """AMP triangle geometry configuration."""
    RECT_WIDTH_CM: float = 13.0
    TRIANGLE_WIDTH_CM: float = 5.5
    TOTAL_WIDTH_CM: float = 18.5
    
    VERTEX_1_X_PCT: float = 0.0
    VERTEX_1_Y_PCT: float = 0.0
    VERTEX_2_X_PCT: float = 0.44
    VERTEX_2_Y_PCT: float = 0.0
    VERTEX_3_X_PCT: float = 0.54
    VERTEX_3_Y_PCT: float = 0.5
    VERTEX_4_X_PCT: float = 0.44
    VERTEX_4_Y_PCT: float = 1.0
    VERTEX_5_X_PCT: float = 0.0
    VERTEX_5_Y_PCT: float = 1.0
    
    EXCLUSION_LEFT_CM: float = 0.0
    EXCLUSION_RIGHT_CM: float = 16.0
    
    COLOR: str = "#0067B9"
    
    def get_vertices_pixels(self, width: int, height: int) -> List[Tuple[int, int]]:
        """Get polygon vertices in pixels."""
        return [
            (int(width * self.VERTEX_1_X_PCT), int(height * self.VERTEX_1_Y_PCT)),
            (int(width * self.VERTEX_2_X_PCT), int(height * self.VERTEX_2_Y_PCT)),
            (int(width * self.VERTEX_3_X_PCT), int(height * self.VERTEX_3_Y_PCT)),
            (int(width * self.VERTEX_4_X_PCT), int(height * self.VERTEX_4_Y_PCT)),
            (int(width * self.VERTEX_5_X_PCT), int(height * self.VERTEX_5_Y_PCT)),
        ]


AMP_CONFIG = AMPTriangleConfig()


# =============================================================================
# FOOTER TEMPLATES
# =============================================================================

FOOTER_TEMPLATES = {
    Confidentiality.PUBLIC: 
        "©{year} Analog Devices, Inc. All Rights Reserved.",
    Confidentiality.CONFIDENTIAL: 
        "Analog Devices Confidential Information. ©{year} Analog Devices, Inc. All Rights Reserved.",
    Confidentiality.INTERNAL_ONLY: 
        "Analog Devices Confidential Information—Not for External Distribution. ©{year} Analog Devices, Inc. All Rights Reserved.",
}


# =============================================================================
# TITLE CASE CONFIGURATION
# =============================================================================

TITLE_CASE_LOWERCASE_WORDS = frozenset({
    'a', 'an', 'the', 'and', 'but', 'or', 'nor',
    'for', 'yet', 'so', 'at', 'by', 'in', 'of',
    'on', 'to', 'up', 'as', 'is', 'if'
})


# =============================================================================
# CONSTANTS
# =============================================================================

CLOSING_TAGLINE = "AHEAD OF WHAT'S POSSIBLE"
COMPANY_URL = "analog.com"


# =============================================================================
# TABLE CONFIGURATION
# =============================================================================

@dataclass
class TableConfig:
    """Table styling configuration."""
    MIN_ROW_HEIGHT_CM: float = 1.0
    DEFAULT_FONT_SIZE: int = 18
    RECOMMENDED_FONT_SIZE: int = 14
    HEADER_BG_COLOR: str = "#B8D4E8"
    ALT_ROW_COLOR: str = "#E8F1F8"


TABLE_CONFIG = TableConfig()


# =============================================================================
# ASSET PATHS CONFIGURATION
# =============================================================================

@dataclass
class AssetPathsConfig:
    """Default asset file paths."""
    BASE_DIR: str = "assets"
    LOGO_BLUE: str = "adi_logo_blue.png"
    LOGO_WHITE: str = "adi_logo_white.png"
    COVER_BACKGROUND: str = "cover_background.jpg"
    TESSELLATED_BACKGROUND: str = "tessellated_bg.jpg"
    AMP_OVERLAY: str = "amp_overlay.png"


ASSET_PATHS = AssetPathsConfig()


# =============================================================================
# VALIDATION RULES
# =============================================================================

@dataclass
class ValidationRules:
    """Brand compliance validation rules."""
    MAX_RECOMMENDED_SLIDES: int = 10
    MIN_TITLE_SIZE_PT: int = 36
    MAX_TITLE_LINES: int = 3
    MIN_TABLE_ROW_HEIGHT_CM: float = 1.0
    MIN_CONTRAST_RATIO: float = 3.0
    BODY_CONTRAST_RATIO: float = 4.5


VALIDATION = ValidationRules()
