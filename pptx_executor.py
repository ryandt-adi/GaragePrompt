#!/usr/bin/env python3
"""
pptx_executor.py
================
Safe execution environment for PowerPoint generation scripts.
Similar to excel_processor.py but for python-pptx.

Version: 1.0.0
"""

import traceback
from typing import Dict, Any, Optional, Tuple, List
from pptx import Presentation

# Import all ADI modules
from adi_template_config import *
from adi_pptx_generator import (
    ADIPresentation, ContentItem, ChartSeries, TableData,
    to_title_case, format_bullet, hex_to_rgb,
    create_presentation, AssetManager, BrandValidator
)

# Import pptx components
from pptx.util import Pt, Cm, Inches, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData

from datetime import datetime, date


class ExecutionResult:
    """Result of script execution."""
    
    def __init__(self, success: bool, presentation: Optional[Any],
                 error_message: str, traceback_str: str):
        self.success = success
        self.presentation = presentation
        self.error_message = error_message
        self.traceback = traceback_str
    
    @property
    def has_error(self) -> bool:
        return not self.success


class PPTXScriptExecutor:
    """
    Executes Python scripts that generate PowerPoint presentations.
    Provides safe execution with ADI template compliance built-in.
    """
    
    def __init__(self):
        self._namespace = self._build_namespace()
    
    def _build_namespace(self) -> Dict[str, Any]:
        """Build execution namespace with all allowed imports."""
        return {
            # Core pptx
            'Presentation': Presentation,
            
            # ADI Template Classes
            'ADIPresentation': ADIPresentation,
            'ContentItem': ContentItem,
            'ChartSeries': ChartSeries,
            'TableData': TableData,
            'create_presentation': create_presentation,
            
            # Configuration
            'COLORS': COLORS,
            'DIMENSIONS': DIMENSIONS,
            'TYPOGRAPHY': TYPOGRAPHY,
            'CONTAINERS': CONTAINERS,
            'AMP_CONFIG': AMP_CONFIG,
            'TABLE_CONFIG': TABLE_CONFIG,
            'VALIDATION': VALIDATION,
            
            # Enums
            'Confidentiality': Confidentiality,
            'SlideType': SlideType,
            'TableStyle': TableStyle,
            'SectionSlideType': SectionSlideType,
            'ChartColorScheme': ChartColorScheme,
            
            # Constants
            'FOOTER_TEMPLATES': FOOTER_TEMPLATES,
            'CLOSING_TAGLINE': CLOSING_TAGLINE,
            'COMPANY_URL': COMPANY_URL,
            
            # Utilities
            'to_title_case': to_title_case,
            'format_bullet': format_bullet,
            'hex_to_rgb': hex_to_rgb,
            'Pt': Pt,
            'Cm': Cm,
            'Inches': Inches,
            'Emu': Emu,
            'RGBColor': RGBColor,
            
            # Enums from pptx
            'PP_ALIGN': PP_ALIGN,
            'MSO_ANCHOR': MSO_ANCHOR,
            'MSO_SHAPE': MSO_SHAPE,
            'XL_CHART_TYPE': XL_CHART_TYPE,
            'XL_LEGEND_POSITION': XL_LEGEND_POSITION,
            
            # Chart
            'CategoryChartData': CategoryChartData,
            
            # Date/time
            'datetime': datetime,
            'date': date,
        }
    
    def validate_script(self, script: str) -> Tuple[bool, str]:
        """Validate script before execution."""
        if not script or not script.strip():
            return False, "Script is empty"
        
        dangerous = [
            ("import os", "Direct os import not allowed"),
            ("import sys", "Direct sys import not allowed"),
            ("import subprocess", "subprocess not allowed"),
            ("__import__", "Dynamic imports not allowed"),
            ("eval(", "eval() not allowed"),
            ("exec(", "Nested exec() not allowed"),
            ("open(", "File operations not allowed"),
            (".save(", "save() should not be included - app handles saving"),
        ]
        
        for pattern, message in dangerous:
            if pattern in script:
                return False, message
        
        if "deck" not in script and "ADIPresentation" not in script:
            return False, "Script must create 'deck' using ADIPresentation"
        
        return True, ""
    
    def execute(self, script: str, validate: bool = True) -> ExecutionResult:
        """Execute script and return generated presentation."""
        if validate:
            is_valid, error = self.validate_script(script)
            if not is_valid:
                return ExecutionResult(False, None, error, "")
        
        namespace = self._namespace.copy()
        
        try:
            exec(script, namespace, namespace)
            
            if 'deck' in namespace:
                pres = namespace['deck']
                if isinstance(pres, ADIPresentation):
                    return ExecutionResult(True, pres, "", "")
                else:
                    return ExecutionResult(False, None, 
                        "'deck' is not an ADIPresentation object", "")
            else:
                return ExecutionResult(False, None,
                    "Script did not create 'deck' object", "")
                
        except SyntaxError as e:
            return ExecutionResult(False, None,
                f"Syntax error at line {e.lineno}: {e.msg}",
                traceback.format_exc())
        except Exception as e:
            return ExecutionResult(False, None, str(e),
                traceback.format_exc())
    
    def save_presentation(self, presentation: ADIPresentation, 
                          filepath: str) -> Tuple[bool, str]:
        """Save presentation to file."""
        try:
            presentation.save(filepath, validate=True)
            return True, ""
        except Exception as e:
            return False, str(e)
    
    def get_available_features(self) -> Dict[str, List[str]]:
        """Get summary of available features."""
        return {
            "slide_types": [st.value for st in SlideType],
            "table_styles": [ts.value for ts in TableStyle],
            "confidentiality_levels": [c.value for c in Confidentiality],
            "chart_types": ["bar", "column", "line", "pie", "area"],
            "helper_classes": [
                "ADIPresentation", "ContentItem", "ChartSeries",
                "TableData", "BrandValidator"
            ],
            "config_modules": [
                "COLORS", "DIMENSIONS", "TYPOGRAPHY", "CONTAINERS",
                "AMP_CONFIG", "TABLE_CONFIG", "VALIDATION"
            ]
        }


# Service instance
pptx_executor = PPTXScriptExecutor()
