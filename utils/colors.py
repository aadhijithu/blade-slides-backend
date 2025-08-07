"""
Color conversion utilities for Figma to PowerPoint
"""

import re

def hex_to_rgb(hex_color):
    """
    Convert hex color to RGB tuple
    
    Args:
        hex_color: Hex color string (e.g., '#FF0000' or 'FF0000')
        
    Returns:
        Tuple of (r, g, b) values (0-255)
    """
    if not hex_color:
        return (0, 0, 0)
    
    # Remove # if present
    hex_color = hex_color.lstrip('#')
    
    # Handle 3-digit hex colors
    if len(hex_color) == 3:
        hex_color = ''.join([c*2 for c in hex_color])
    
    # Validate hex format
    if not re.match(r'^[0-9A-Fa-f]{6}$', hex_color):
        print(f"Invalid hex color: {hex_color}, using black")
        return (0, 0, 0)
    
    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return (r, g, b)
    except ValueError:
        print(f"Error converting hex color: {hex_color}, using black")
        return (0, 0, 0)

def rgb_to_hex(r, g, b):
    """
    Convert RGB values to hex color string
    
    Args:
        r, g, b: RGB values (0-255)
        
    Returns:
        Hex color string with # prefix
    """
    return f"#{r:02x}{g:02x}{b:02x}".upper()

def figma_rgb_to_hex(figma_rgb):
    """
    Convert Figma RGB object (0-1 range) to hex color
    
    Args:
        figma_rgb: Dictionary with 'r', 'g', 'b' keys (0-1 range)
        
    Returns:
        Hex color string
    """
    if not figma_rgb or not isinstance(figma_rgb, dict):
        return '#000000'
    
    try:
        r = int(figma_rgb.get('r', 0) * 255)
        g = int(figma_rgb.get('g', 0) * 255)
        b = int(figma_rgb.get('b', 0) * 255)
        return rgb_to_hex(r, g, b)
    except (TypeError, ValueError):
        return '#000000'

def normalize_color(color_value):
    """
    Normalize various color formats to hex
    
    Args:
        color_value: Color in various formats (hex, rgb object, etc.)
        
    Returns:
        Normalized hex color string
    """
    if not color_value:
        return '#000000'
    
    # Already hex
    if isinstance(color_value, str):
        if color_value.startswith('#'):
            return color_value
        elif color_value in ['transparent', 'none']:
            return None
        else:
            # Try to parse as hex without #
            return f"#{color_value}" if re.match(r'^[0-9A-Fa-f]{6}$', color_value) else '#000000'
    
    # Figma RGB object
    if isinstance(color_value, dict) and 'r' in color_value:
        return figma_rgb_to_hex(color_value)
    
    return '#000000'

def is_transparent_color(color_value):
    """
    Check if color should be treated as transparent
    
    Args:
        color_value: Color value in any format
        
    Returns:
        Boolean indicating if color is transparent
    """
    if not color_value:
        return True
    
    if isinstance(color_value, str):
        return color_value.lower() in ['transparent', 'none', '']
    
    return False

def get_contrast_color(hex_color):
    """
    Get contrasting color (black or white) for given color
    
    Args:
        hex_color: Hex color string
        
    Returns:
        '#000000' for light colors, '#FFFFFF' for dark colors
    """
    rgb = hex_to_rgb(hex_color)
    
    # Calculate luminance
    luminance = (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) / 255
    
    return '#000000' if luminance > 0.5 else '#FFFFFF'

def lighten_color(hex_color, factor=0.1):
    """
    Lighten a color by a given factor
    
    Args:
        hex_color: Hex color string
        factor: Lightening factor (0-1)
        
    Returns:
        Lightened hex color string
    """
    rgb = hex_to_rgb(hex_color)
    
    r = min(255, int(rgb[0] + (255 - rgb[0]) * factor))
    g = min(255, int(rgb[1] + (255 - rgb[1]) * factor))
    b = min(255, int(rgb[2] + (255 - rgb[2]) * factor))
    
    return rgb_to_hex(r, g, b)

def darken_color(hex_color, factor=0.1):
    """
    Darken a color by a given factor
    
    Args:
        hex_color: Hex color string
        factor: Darkening factor (0-1)
        
    Returns:
        Darkened hex color string
    """
    rgb = hex_to_rgb(hex_color)
    
    r = max(0, int(rgb[0] * (1 - factor)))
    g = max(0, int(rgb[1] * (1 - factor)))
    b = max(0, int(rgb[2] * (1 - factor)))
    
    return rgb_to_hex(r, g, b) 