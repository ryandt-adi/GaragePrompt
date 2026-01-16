"""
config.py
Centralized configuration for Analog Garage Workbench.
All constants, colors, paths, and application settings.
"""

import os
import sys

# =============================================================================
# APPLICATION METADATA
# =============================================================================

APP_NAME = "Analog Garage Workbench"
APP_VERSION = "3.0"
APP_AUTHOR = "Analog Devices, Inc."
APP_CONTACT = "david.ryan@analog.com"
APP_DESCRIPTION = "Multi-Output Prompt Generator + Script Executor"

# =============================================================================
# OUTPUT TYPES CONFIGURATION
# =============================================================================

OUTPUT_TYPES = {
    "excel_value_model": {
        "id": "excel_value_model",
        "name": "Excel Value Creation Model",
        "short_name": "Excel Model",
        "description": "Comprehensive Excel financial model with charts, projections, and sensitivity analysis",
        "output_format": "xlsx",
        "icon": "📊",
        "template_id": "value_creation_v3",
        "script_type": "excel",
        "color": "#28A745",  # Green
    },
    "powerpoint_pitch": {
        "id": "powerpoint_pitch",
        "name": "Executive Pitch Deck",
        "short_name": "PowerPoint Deck",
        "description": "Executive-ready PowerPoint presentation with key insights and visualizations",
        "output_format": "pptx",
        "icon": "📽️",
        "template_id": "executive_pitch_v1",
        "script_type": "powerpoint",
        "color": "#FF6600",  # Orange
    },
    "word_gonogo": {
        "id": "word_gonogo",
        "name": "Deep Dive Summary - Recommend to Explore",
        "short_name": "Deep Dive Summary",
        "description": "Comprehensive Deep Dive assessment with exploration recommendation for Analog Garage",
        "output_format": "docx",
        "icon": "📋",
        "template_id": "gonogo_report_v1",
        "script_type": "word",
        "color": "#0067B9",  # Blue
    },
}


# List for dropdown
OUTPUT_TYPE_OPTIONS = [
    f"{v['icon']} {v['name']}" for v in OUTPUT_TYPES.values()
]

# Mapping from display name to ID
OUTPUT_TYPE_DISPLAY_TO_ID = {
    f"{v['icon']} {v['name']}": k for k, v in OUTPUT_TYPES.items()
}

# =============================================================================
# BRAND COLORS
# =============================================================================

ADI_COLORS = {
    # Primary palette
    "primary_blue": "#0067B9",
    "dark_blue": "#003D6A",
    "light_blue": "#4A9BD9",
    "accent_orange": "#FF6600",
    
    # Neutrals
    "white": "#FFFFFF",
    "light_gray": "#F5F5F5",
    "medium_gray": "#E0E0E0",
    "dark_gray": "#333333",
    "text_gray": "#666666",
    
    # Code editor
    "code_bg": "#1E2433",
    "code_fg": "#E8E8E8",
    
    # Status colors
    "success_green": "#28A745",
    "error_red": "#DC3545",
    "warning_yellow": "#FFC107",
}

# =============================================================================
# UI CONFIGURATION
# =============================================================================

UI_CONFIG = {
    "window_width": 1100,
    "window_height": 900,
    "font_family": "Arial",
    "code_font_family": "Courier New",
    "header_font_size": 24,
    "title_font_size": 14,
    "body_font_size": 11,
    "small_font_size": 10,
}

# =============================================================================
# FORM DEFAULTS
# =============================================================================

FORM_DEFAULTS = {
    "geographic_scope": "Global",
    "analysis_timeframe": "Year 1 at Scale",
    "innovation_stage": "Concept",
    "currency": "USD",
    "output_type": "excel_value_model",
}

FORM_OPTIONS = {
    "geographic_scope": [
        "United States",
        "North America",
        "Europe",
        "Asia-Pacific",
        "Global",
        "Other",  # ✅ ADDED
    ],
    "analysis_timeframe": [
        "Year 1 at Scale",
        "3-Year Projection",
        "5-Year Projection",
        "10-Year Projection",
        "Other",  # ✅ ADDED
    ],
    "innovation_stage": [
        "Concept",
        "Prototype",
        "Pilot",
        "Commercial",
        "Growth",
        "Other",  # ✅ ADDED
    ],
    "currency": [
        "USD", 
        "EUR", 
        "GBP", 
        "JPY",
        "Other",  # ✅ ADDED
    ],
    "output_type": OUTPUT_TYPE_OPTIONS,
}

# =============================================================================
# FILE PATHS
# =============================================================================

def get_script_directory():
    """Get the directory where the script is located."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

SCRIPT_DIR = get_script_directory()
LOGO_PATH = os.path.join(SCRIPT_DIR, "adi_logo.png")

# =============================================================================
# EXCEL EXECUTION NAMESPACE IMPORTS
# =============================================================================

EXCEL_ALLOWED_IMPORTS = [
    "openpyxl",
    "Font",
    "PatternFill",
    "Alignment",
    "Border",
    "Side",
    "get_column_letter",
    "LineChart",
    "BarChart",
    "Reference",
    "DataValidation",
    "datetime",
    "relativedelta",
]

# =============================================================================
# FEATURE FLAGS
# =============================================================================

FEATURES = {
    "enable_logo": True,
    "enable_session_logging": True,
    "enable_export_log": True,
    "demo_mode": True,
    "enable_powerpoint": True,
    "enable_word": True,
}

# =============================================================================
# SCRIPT INSTRUCTIONS BY OUTPUT TYPE
# =============================================================================

SCRIPT_INSTRUCTIONS = {
    "excel": {
        "title": "Excel Script Input",
        "instruction": "💡 Paste the AI-generated Python script below. The script must create a workbook object named 'wb' using openpyxl. Do NOT include wb.save().",
        "button_text": "▶️ Execute & Save Excel",
        "file_extension": ".xlsx",
        "file_type": "Excel files",
    },
    "powerpoint": {
        "title": "PowerPoint Script Input",
        "instruction": "💡 Paste the AI-generated Python script below. The script must create a presentation object named 'prs' using python-pptx. Do NOT include prs.save().",
        "button_text": "▶️ Execute & Save PowerPoint",
        "file_extension": ".pptx",
        "file_type": "PowerPoint files",
    },
    "word": {
        "title": "Word Document Script Input",
        "instruction": "💡 Paste the AI-generated Python script below. The script must create a document object named 'doc' using python-docx. Do NOT include doc.save().",
        "button_text": "▶️ Execute & Save Word Doc",
        "file_extension": ".docx",
        "file_type": "Word files",
    },
}


# =============================================================================
# OUTPUT TYPES CONFIGURATION
# =============================================================================

OUTPUT_TYPES = {
    "excel_value_model": {
        "id": "excel_value_model",
        "name": "Excel Value Creation Model",
        "short_name": "Excel Model",
        "description": "Comprehensive Excel financial model with charts, projections, and sensitivity analysis",
        "output_format": "xlsx",
        "icon": "📊",
        "template_id": "value_creation_v3",
        "script_type": "excel",
        "color": "#28A745",
    },
    "powerpoint_pitch": {
        "id": "powerpoint_pitch",
        "name": "Executive Pitch Deck",
        "short_name": "PowerPoint Deck",
        "description": "Executive-ready PowerPoint presentation with key insights and visualizations",
        "output_format": "pptx",
        "icon": "📽️",
        "template_id": "executive_pitch_v1",
        "script_type": "powerpoint",
        "color": "#FF6600",
    },
    "word_gonogo": {
        "id": "word_gonogo",
        "name": "Deep Dive Summary - Recommend to Explore",
        "short_name": "Deep Dive Summary",
        "description": "Comprehensive Deep Dive assessment with exploration recommendation for Analog Garage",
        "output_format": "docx",
        "icon": "📋",
        "template_id": "gonogo_report_v1",
        "script_type": "word",
        "color": "#0067B9",
    },
}

# List for dropdown
OUTPUT_TYPE_OPTIONS = [
    f"{v['icon']} {v['name']}" for v in OUTPUT_TYPES.values()
]

# Mapping from display name to ID
OUTPUT_TYPE_DISPLAY_TO_ID = {
    f"{v['icon']} {v['name']}": k for k, v in OUTPUT_TYPES.items()
}

# Add to FORM_OPTIONS
FORM_OPTIONS["output_type"] = OUTPUT_TYPE_OPTIONS

# =============================================================================
# SCRIPT INSTRUCTIONS BY OUTPUT TYPE
# =============================================================================

SCRIPT_INSTRUCTIONS = {
    "excel": {
        "title": "Excel Script Input",
        "instruction": "💡 Paste the AI-generated Python script below. The script must create a workbook object named 'wb' using openpyxl. Do NOT include wb.save().",
        "button_text": "▶️ Execute & Save Excel",
        "file_extension": ".xlsx",
        "file_type": "Excel files",
    },
    "powerpoint": {
        "title": "PowerPoint Script Input",
        "instruction": "💡 Paste the AI-generated Python script below. The script must create a presentation object named 'prs' using python-pptx. Do NOT include prs.save().",
        "button_text": "▶️ Execute & Save PowerPoint",
        "file_extension": ".pptx",
        "file_type": "PowerPoint files",
    },
    "word": {
        "title": "Word Document Script Input",
        "instruction": "💡 Paste the AI-generated Python script below. The script must create a document object named 'doc' using python-docx. Do NOT include doc.save().",
        "button_text": "▶️ Execute & Save Word Doc",
        "file_extension": ".docx",
        "file_type": "Word files",
    },
}