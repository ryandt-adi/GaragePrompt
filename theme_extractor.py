#!/usr/bin/env python3
"""
theme_extractor.py
==================
Extract PowerPoint theme colors from a template and apply them to generated presentations.

Usage:
    One-time extraction:
        python theme_extractor.py template.pptx
    
    This creates theme_config.py with all extracted colors that can be imported
    into other modules.

Version: 1.0.0
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, Optional, List, Tuple
from datetime import datetime
import json
import shutil
import tempfile
import os

# Office Open XML namespaces
OOXML_NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}

# Register namespaces to preserve them when writing
for prefix, uri in OOXML_NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class ThemeColorExtractor:
    """
    Extract theme colors from a PowerPoint template file.
    
    PowerPoint theme colors include:
    - dk1, lt1: Primary dark/light (usually black/white)
    - dk2, lt2: Secondary dark/light
    - accent1-6: Six accent colors
    - hlink: Hyperlink color
    - folHlink: Followed hyperlink color
    """
    
    # Standard theme color names in order
    COLOR_NAMES = [
        'dk1', 'lt1', 'dk2', 'lt2',
        'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
        'hlink', 'folHlink'
    ]
    
    def __init__(self, template_path: str):
        """Initialize with path to template.pptx"""
        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")
        
        self.theme_colors: Dict[str, str] = {}
        self.theme_name: str = ""
        self.font_scheme: Dict[str, str] = {}
        self.raw_theme_xml: Optional[bytes] = None
    
    def extract(self) -> Dict[str, str]:
        """
        Extract all theme colors from the template.
        
        Returns:
            Dictionary mapping color names to hex values (without #)
        """
        with zipfile.ZipFile(self.template_path, 'r') as zf:
            # Find theme file (usually ppt/theme/theme1.xml)
            theme_path = self._find_theme_path(zf)
            
            if not theme_path:
                raise FileNotFoundError("No theme file found in template")
            
            # Store raw XML for potential direct copying
            self.raw_theme_xml = zf.read(theme_path)
            
            # Parse theme XML
            root = ET.fromstring(self.raw_theme_xml)
            
            # Extract theme name
            self.theme_name = root.get('name', 'Unknown Theme')
            
            # Find and parse color scheme
            self._parse_theme_elements(root)
        
        return self.theme_colors
    
    def _find_theme_path(self, zf: zipfile.ZipFile) -> Optional[str]:
        """Find the theme XML file within the PPTX archive."""
        # Standard location
        if 'ppt/theme/theme1.xml' in zf.namelist():
            return 'ppt/theme/theme1.xml'
        
        # Search for theme files
        theme_files = [
            f for f in zf.namelist() 
            if 'theme' in f.lower() and f.endswith('.xml')
        ]
        
        return theme_files[0] if theme_files else None
    
    def _parse_theme_elements(self, root: ET.Element):
        """Parse theme elements including colors and fonts."""
        # Namespace-aware tag matching
        for elem in root.iter():
            tag_local = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            
            if tag_local == 'clrScheme':
                self._parse_color_scheme(elem)
            elif tag_local == 'fontScheme':
                self._parse_font_scheme(elem)
    
    def _parse_color_scheme(self, clr_scheme: ET.Element):
        """Parse the color scheme element."""
        for child in clr_scheme:
            tag_local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            
            if tag_local in self.COLOR_NAMES:
                color_value = self._extract_color_value(child)
                if color_value:
                    self.theme_colors[tag_local] = color_value
    
    def _extract_color_value(self, color_elem: ET.Element) -> Optional[str]:
        """
        Extract hex color value from a color element.
        
        Colors can be defined as:
        - srgbClr: Direct RGB hex value
        - sysClr: System color with lastClr attribute
        - schemeClr: Reference to another scheme color
        """
        for child in color_elem:
            tag_local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            
            if tag_local == 'srgbClr':
                return child.get('val', '').upper()
            elif tag_local == 'sysClr':
                # System colors store actual value in lastClr
                return child.get('lastClr', '').upper()
        
        return None
    
    def _parse_font_scheme(self, font_scheme: ET.Element):
        """Parse font scheme for major/minor fonts."""
        for child in font_scheme:
            tag_local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            
            if tag_local in ('majorFont', 'minorFont'):
                for font_child in child:
                    font_tag = font_child.tag.split('}')[-1] if '}' in font_child.tag else font_child.tag
                    if font_tag == 'latin':
                        typeface = font_child.get('typeface', '')
                        if typeface:
                            self.font_scheme[tag_local] = typeface
    
    def get_color_mapping(self) -> Dict[str, str]:
        """
        Get a semantic mapping of theme colors.
        
        Returns:
            Dictionary with semantic names mapped to hex colors
        """
        return {
            'primary_dark': self.theme_colors.get('dk1', '000000'),
            'primary_light': self.theme_colors.get('lt1', 'FFFFFF'),
            'secondary_dark': self.theme_colors.get('dk2', '1F497D'),
            'secondary_light': self.theme_colors.get('lt2', 'EEECE1'),
            'accent1': self.theme_colors.get('accent1', '4F81BD'),
            'accent2': self.theme_colors.get('accent2', 'C0504D'),
            'accent3': self.theme_colors.get('accent3', '9BBB59'),
            'accent4': self.theme_colors.get('accent4', '8064A2'),
            'accent5': self.theme_colors.get('accent5', '4BACC6'),
            'accent6': self.theme_colors.get('accent6', 'F79646'),
            'hyperlink': self.theme_colors.get('hlink', '0000FF'),
            'followed_hyperlink': self.theme_colors.get('folHlink', '800080'),
        }
    
    def generate_python_config(self, output_path: str = 'theme_config.py') -> str:
        """
        Generate a Python configuration file with extracted theme colors.
        
        Args:
            output_path: Path for the generated Python file
            
        Returns:
            Path to the generated file
        """
        if not self.theme_colors:
            self.extract()
        
        mapping = self.get_color_mapping()
        
        content = f'''#!/usr/bin/env python3
"""
theme_config.py
===============
Auto-generated PowerPoint theme colors extracted from template.

Source Template: {self.template_path.name}
Theme Name: {self.theme_name}
Extracted: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

DO NOT EDIT MANUALLY - Regenerate using:
    python theme_extractor.py {self.template_path.name}
"""

from typing import Dict, Tuple

# =============================================================================
# RAW THEME COLORS (as extracted from template)
# =============================================================================

THEME_COLORS_RAW: Dict[str, str] = {{
{self._format_dict_items(self.theme_colors, indent=4)}
}}

# =============================================================================
# SEMANTIC COLOR MAPPING
# =============================================================================

THEME_COLORS: Dict[str, str] = {{
{self._format_dict_items(mapping, indent=4)}
}}

# =============================================================================
# ACCENT COLOR LIST (for charts and data visualization)
# =============================================================================

ACCENT_COLORS: Tuple[str, ...] = (
    "#{mapping['accent1']}",
    "#{mapping['accent2']}",
    "#{mapping['accent3']}",
    "#{mapping['accent4']}",
    "#{mapping['accent5']}",
    "#{mapping['accent6']}",
)

# =============================================================================
# FONT SCHEME
# =============================================================================

THEME_FONTS: Dict[str, str] = {{
{self._format_dict_items(self.font_scheme, indent=4)}
}}

# =============================================================================
# CONVENIENCE CONSTANTS
# =============================================================================

# Primary brand color (accent1)
PRIMARY_COLOR = "#{mapping['accent1']}"

# Background colors
DARK_BACKGROUND = "#{mapping['primary_dark']}"
LIGHT_BACKGROUND = "#{mapping['primary_light']}"

# Text colors
TEXT_ON_LIGHT = "#{mapping['primary_dark']}"
TEXT_ON_DARK = "#{mapping['primary_light']}"

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
    return f"{{r:02X}}{{g:02X}}{{b:02X}}"
'''
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        return output_path
    
    def generate_json_config(self, output_path: str = 'theme_colors.json') -> str:
        """Generate JSON configuration file."""
        if not self.theme_colors:
            self.extract()
        
        config = {
            'source_template': str(self.template_path),
            'theme_name': self.theme_name,
            'extracted_date': datetime.now().isoformat(),
            'raw_colors': self.theme_colors,
            'semantic_mapping': self.get_color_mapping(),
            'fonts': self.font_scheme,
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2)
        
        return output_path
    
    def _format_dict_items(self, d: Dict[str, str], indent: int = 4) -> str:
        """Format dictionary items for Python code generation."""
        lines = []
        for key, value in d.items():
            lines.append(f'{" " * indent}"{key}": "{value}",')
        return '\n'.join(lines)


class ThemeApplier:
    """
    Apply theme colors to a PowerPoint presentation.
    
    Supports two methods:
    1. Copy entire theme from template (most reliable)
    2. Modify individual color values in existing theme
    """
    
    def __init__(self, config_path: str = 'theme_colors.json'):
        """Initialize with path to theme configuration."""
        self.config_path = Path(config_path)
        
        if config_path.endswith('.json'):
            with open(config_path, 'r') as f:
                config = json.load(f)
            self.theme_colors = config.get('raw_colors', {})
        else:
            # Assume Python module - import dynamically
            import importlib.util
            spec = importlib.util.spec_from_file_location("theme_config", config_path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            self.theme_colors = getattr(module, 'THEME_COLORS_RAW', {})
    
    @staticmethod
    def copy_theme_from_template(
        template_path: str, 
        target_pptx_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """
        Copy the entire theme from a template to a target presentation.
        
        This is the most reliable method to ensure complete theme consistency.
        
        Args:
            template_path: Path to source template.pptx
            target_pptx_path: Path to target presentation
            output_path: Optional output path (defaults to overwriting target)
            
        Returns:
            Path to the output file
        """
        output_path = output_path or target_pptx_path
        
        # Create temp directory for extraction
        with tempfile.TemporaryDirectory() as temp_dir:
            template_dir = Path(temp_dir) / 'template'
            target_dir = Path(temp_dir) / 'target'
            
            # Extract both files
            with zipfile.ZipFile(template_path, 'r') as zf:
                zf.extractall(template_dir)
            
            with zipfile.ZipFile(target_pptx_path, 'r') as zf:
                zf.extractall(target_dir)
            
            # Copy theme folder from template to target
            template_theme = template_dir / 'ppt' / 'theme'
            target_theme = target_dir / 'ppt' / 'theme'
            
            if template_theme.exists():
                # Remove existing theme
                if target_theme.exists():
                    shutil.rmtree(target_theme)
                
                # Copy template theme
                shutil.copytree(template_theme, target_theme)
            
            # Also copy slide layouts and masters for complete consistency
            for folder in ['slideMasters', 'slideLayouts']:
                template_folder = template_dir / 'ppt' / folder
                target_folder = target_dir / 'ppt' / folder
                
                if template_folder.exists():
                    if target_folder.exists():
                        shutil.rmtree(target_folder)
                    shutil.copytree(template_folder, target_folder)
            
            # Repackage as PPTX
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, dirs, files in os.walk(target_dir):
                    for file in files:
                        file_path = Path(root) / file
                        arcname = file_path.relative_to(target_dir)
                        zf.write(file_path, arcname)
        
        return output_path
    
    def apply_colors_to_theme_xml(
        self, 
        pptx_path: str, 
        output_path: Optional[str] = None
    ) -> str:
        """
        Apply theme colors by modifying the theme XML directly.
        
        This method updates color values while preserving other theme elements.
        
        Args:
            pptx_path: Path to the presentation
            output_path: Optional output path
            
        Returns:
            Path to the output file
        """
        output_path = output_path or pptx_path
        
        with tempfile.TemporaryDirectory() as temp_dir:
            extract_dir = Path(temp_dir) / 'pptx'
            
            # Extract PPTX
            with zipfile.ZipFile(pptx_path, 'r') as zf:
                zf.extractall(extract_dir)
            
            # Find and modify theme file
            theme_path = extract_dir / 'ppt' / 'theme' / 'theme1.xml'
            
            if theme_path.exists():
                self._update_theme_xml(theme_path)
            
            # Repackage
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        file_path = Path(root) / file
                        arcname = file_path.relative_to(extract_dir)
                        zf.write(file_path, arcname)
        
        return output_path
    
    def _update_theme_xml(self, theme_path: Path):
        """Update color values in theme XML file."""
        tree = ET.parse(theme_path)
        root = tree.getroot()
        
        # Find color scheme
        for elem in root.iter():
            tag_local = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            
            if tag_local == 'clrScheme':
                self._update_color_scheme(elem)
                break
        
        # Write back
        tree.write(theme_path, xml_declaration=True, encoding='UTF-8')
    
    def _update_color_scheme(self, clr_scheme: ET.Element):
        """Update colors in the color scheme element."""
        a_ns = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
        
        for color_name, hex_value in self.theme_colors.items():
            if not hex_value:
                continue
            
            for child in clr_scheme:
                tag_local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                
                if tag_local == color_name:
                    # Clear existing color definition
                    for subchild in list(child):
                        child.remove(subchild)
                    
                    # Add new srgbClr element
                    new_color = ET.SubElement(child, f'{a_ns}srgbClr')
                    new_color.set('val', hex_value)
                    break


# =============================================================================
# INTEGRATION WITH ADI PRESENTATION GENERATOR
# =============================================================================

def integrate_with_adi_generator(config_path: str = 'theme_config.py'):
    """
    Generate integration code for adi_template_config.py.
    
    Call this after extraction to see how to update the ADI config.
    """
    print("""
# =============================================================================
# Add to adi_template_config.py to use extracted theme colors
# =============================================================================

# Option 1: Import generated config
try:
    from theme_config import (
        THEME_COLORS, ACCENT_COLORS, PRIMARY_COLOR,
        DARK_BACKGROUND, LIGHT_BACKGROUND
    )
    USE_EXTRACTED_THEME = True
except ImportError:
    USE_EXTRACTED_THEME = False

# Option 2: Directly update ADIColorPalette if extraction was successful
if USE_EXTRACTED_THEME:
    @dataclass
    class ADIColorPalette:
        # Primary (from theme accent1)
        PRIMARY_BLUE: str = PRIMARY_COLOR.lstrip('#')
        
        # Backgrounds (from theme dk1/lt1)
        DARK_BLUE: str = DARK_BACKGROUND.lstrip('#')
        WHITE: str = LIGHT_BACKGROUND.lstrip('#')
        
        # Accents (from theme accent1-6)
        CHART_LIGHT_1: str = ACCENT_COLORS[0].lstrip('#')
        CHART_LIGHT_2: str = ACCENT_COLORS[1].lstrip('#')
        CHART_LIGHT_3: str = ACCENT_COLORS[2].lstrip('#')
        CHART_LIGHT_4: str = ACCENT_COLORS[3].lstrip('#')
        CHART_LIGHT_5: str = ACCENT_COLORS[4].lstrip('#')
        CHART_LIGHT_6: str = ACCENT_COLORS[5].lstrip('#')
        
        # ... rest of the class
""")


# =============================================================================
# COMMAND LINE INTERFACE
# =============================================================================

def main():
    """Command line interface for theme extraction."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Extract PowerPoint theme colors from a template',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
    # Extract theme and generate Python config
    python theme_extractor.py template.pptx
    
    # Extract to specific output files
    python theme_extractor.py template.pptx --python theme_config.py --json theme.json
    
    # Copy theme from template to another presentation
    python theme_extractor.py template.pptx --apply-to presentation.pptx
        '''
    )
    
    parser.add_argument('template', help='Path to template.pptx')
    parser.add_argument('--python', '-p', default='theme_config.py',
                        help='Output Python config file (default: theme_config.py)')
    parser.add_argument('--json', '-j', default='theme_colors.json',
                        help='Output JSON config file (default: theme_colors.json)')
    parser.add_argument('--apply-to', '-a', metavar='PPTX',
                        help='Apply extracted theme to another presentation')
    parser.add_argument('--show-integration', '-i', action='store_true',
                        help='Show integration code for adi_template_config.py')
    
    args = parser.parse_args()
    
    # Extract theme
    print(f"📊 Extracting theme from: {args.template}")
    extractor = ThemeColorExtractor(args.template)
    colors = extractor.extract()
    
    # Display extracted colors
    print(f"\n🎨 Theme Name: {extractor.theme_name}")
    print("\n📋 Extracted Colors:")
    print("-" * 40)
    
    for name, value in colors.items():
        # Create color preview (ANSI escape for terminal)
        preview = f"\033[48;2;{int(value[0:2], 16)};{int(value[2:4], 16)};{int(value[4:6], 16)}m  \033[0m"
        print(f"  {name:12} #{value}  {preview}")
    
    # Generate config files
    print("\n📁 Generating config files:")
    
    py_path = extractor.generate_python_config(args.python)
    print(f"  ✓ Python: {py_path}")
    
    json_path = extractor.generate_json_config(args.json)
    print(f"  ✓ JSON:   {json_path}")
    
    # Apply to another presentation if requested
    if args.apply_to:
        print(f"\n🔄 Applying theme to: {args.apply_to}")
        output = ThemeApplier.copy_theme_from_template(
            args.template, 
            args.apply_to
        )
        print(f"  ✓ Theme applied: {output}")
    
    # Show integration code if requested
    if args.show_integration:
        integrate_with_adi_generator()
    
    print("\n✅ Done!")
    print(f"\nTo use in your code:")
    print(f"    from theme_config import THEME_COLORS, ACCENT_COLORS")


if __name__ == '__main__':
    main()
