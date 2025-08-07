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

def map_figma_font(figma_font):
    """
    Map Figma font to PowerPoint-compatible font
    
    Args:
        figma_font: Font family name from Figma
        
    Returns:
        PowerPoint-compatible font name
    """
    if not figma_font:
        return 'Arial'
    
    # Try exact match first
    if figma_font in FONT_MAPPING:
        return FONT_MAPPING[figma_font]
    
    # Try partial match for font families with weights
    for figma_key, ppt_font in FONT_MAPPING.items():
        if figma_key.lower() in figma_font.lower():
            return ppt_font
    
    # Default fallback
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