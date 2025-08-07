"""
Font mapping and text formatting utilities
"""

from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# Figma to PowerPoint font mapping
FONT_MAPPING = {
    'Inter': 'Calibri',
    'Roboto': 'Arial',
    'Helvetica': 'Arial',
    'SF Pro Display': 'Segoe UI',
    'SF Pro Text': 'Segoe UI',
    'Times': 'Times New Roman',
    'Times New Roman': 'Times New Roman',
    'Georgia': 'Georgia',
    'Courier': 'Courier New',
    'Courier New': 'Courier New',
    'Arial': 'Arial',
    'Calibri': 'Calibri',
    'Segoe UI': 'Segoe UI',
    'Open Sans': 'Calibri',
    'Lato': 'Calibri',
    'Montserrat': 'Calibri',
    'Source Sans Pro': 'Arial'
}

def map_figma_font(figma_font_name):
    """
    Map Figma font names to PowerPoint-compatible font names
    
    Args:
        figma_font_name: Font name from Figma
        
    Returns:
        PowerPoint-compatible font name
    """
    if not figma_font_name:
        return 'Arial'
    
    # Normalize font name (remove extra spaces, convert to lowercase)
    font_lower = figma_font_name.lower().strip()
    
    # Direct mappings for common fonts
    font_mappings = {
        # Inter font family
        'inter': 'Calibri',
        'inter tight': 'Calibri',  # Your specific font
        'inter variable': 'Calibri',
        'inter regular': 'Calibri',
        'inter medium': 'Calibri',
        'inter bold': 'Calibri',
        'inter semi bold': 'Calibri',
        'inter semibold': 'Calibri',
        
        # Google Fonts
        'roboto': 'Calibri',
        'open sans': 'Calibri',
        'lato': 'Calibri',
        'montserrat': 'Calibri',
        'source sans pro': 'Calibri',
        'poppins': 'Calibri',
        'nunito': 'Calibri',
        'work sans': 'Calibri',
        
        # System fonts
        'sf pro': 'Calibri',
        'sf pro display': 'Calibri',
        'sf pro text': 'Calibri',
        'helvetica neue': 'Arial',
        'helvetica': 'Arial',
        'system-ui': 'Calibri',
        '-apple-system': 'Calibri',
        
        # Serif fonts
        'times new roman': 'Times New Roman',
        'times': 'Times New Roman',
        'georgia': 'Georgia',
        'serif': 'Times New Roman',
        
        # Monospace fonts
        'monaco': 'Consolas',
        'menlo': 'Consolas',
        'courier new': 'Courier New',
        'courier': 'Courier New',
        'monospace': 'Consolas',
        
        # Default mappings
        'sans-serif': 'Calibri',
        'arial': 'Arial',
        'calibri': 'Calibri',
        'verdana': 'Verdana',
        'tahoma': 'Tahoma',
        'trebuchet ms': 'Trebuchet MS',
    }
    
    # Check direct mapping first
    if font_lower in font_mappings:
        mapped_font = font_mappings[font_lower]
        print(f"    Font mapping: '{figma_font_name}' -> '{mapped_font}'")
        return mapped_font
    
    # Check partial matches for font families
    for figma_key, ppt_font in font_mappings.items():
        if figma_key in font_lower or font_lower.startswith(figma_key):
            print(f"    Font mapping (partial): '{figma_font_name}' -> '{ppt_font}'")
            return ppt_font
    
    # Fallback to Arial for unknown fonts
    print(f"    Font mapping (fallback): '{figma_font_name}' -> 'Arial'")
    return 'Arial'

def map_text_alignment(figma_align):
    """
    Map Figma text alignment to PowerPoint alignment
    
    Args:
        figma_align: Figma text alignment ('left', 'center', 'right', 'justified')
        
    Returns:
        PowerPoint alignment constant
    """
    alignment_map = {
        'left': PP_ALIGN.LEFT,
        'center': PP_ALIGN.CENTER,
        'right': PP_ALIGN.RIGHT,
        'justified': PP_ALIGN.JUSTIFY,
        'justify': PP_ALIGN.JUSTIFY
    }
    
    return alignment_map.get(figma_align.lower(), PP_ALIGN.LEFT)

def map_vertical_alignment(figma_valign):
    """
    Map Figma vertical alignment to PowerPoint vertical anchor
    
    Args:
        figma_valign: Figma vertical alignment ('top', 'center', 'bottom')
        
    Returns:
        PowerPoint vertical anchor constant
    """
    valign_map = {
        'top': MSO_ANCHOR.TOP,
        'center': MSO_ANCHOR.MIDDLE,
        'middle': MSO_ANCHOR.MIDDLE,
        'bottom': MSO_ANCHOR.BOTTOM
    }
    
    return valign_map.get(figma_valign.lower(), MSO_ANCHOR.TOP)

def get_font_size_multiplier():
    """
    Get font size multiplier for better PowerPoint rendering
    
    Returns:
        Float multiplier for font sizes
    """
    return 1.0  # No scaling needed with python-pptx

def is_bold_weight(font_weight):
    """
    Determine if font weight should be bold in PowerPoint
    
    Args:
        font_weight: Figma font weight string
        
    Returns:
        Boolean indicating if text should be bold
    """
    if not font_weight:
        return False
    
    bold_indicators = ['bold', 'semibold', 'extrabold', '600', '700', '800', '900']
    
    return any(indicator in font_weight.lower() for indicator in bold_indicators)

def is_italic_style(font_weight):
    """
    Determine if font style should be italic in PowerPoint
    
    Args:
        font_weight: Figma font weight/style string
        
    Returns:
        Boolean indicating if text should be italic
    """
    if not font_weight:
        return False
    
    return 'italic' in font_weight.lower() or 'oblique' in font_weight.lower() 