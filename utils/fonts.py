"""
Font mapping and text formatting utilities
"""

from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

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
        # Inter font family - Keep as Inter since it's available in Google Slides
        'inter': 'Inter',
        'inter tight': 'Inter',  # Your specific font - keep as Inter
        'inter variable': 'Inter',
        'inter regular': 'Inter',
        'inter medium': 'Inter',
        'inter bold': 'Inter',
        'inter semi bold': 'Inter',
        'inter semibold': 'Inter',
        
        # TASA ORBIT -> Sora (available in Google Slides)
        'tasa orbit': 'Sora',
        'tasa-orbit': 'Sora',
        'tasaorbit': 'Sora',
        
        # Google Fonts (commonly available)
        'roboto': 'Roboto',
        'open sans': 'Open Sans',
        'lato': 'Lato',
        'montserrat': 'Montserrat',
        'source sans pro': 'Source Sans Pro',
        'poppins': 'Poppins',
        'nunito': 'Nunito',
        'work sans': 'Work Sans',
        'sora': 'Sora',
        
        # System fonts - fallback to Google Slides compatible fonts
        'sf pro': 'Inter',
        'sf pro display': 'Inter',
        'sf pro text': 'Inter',
        'helvetica neue': 'Arial',
        'helvetica': 'Arial',
        'system-ui': 'Inter',
        '-apple-system': 'Inter',
        
        # Serif fonts
        'times new roman': 'Times New Roman',
        'times': 'Times New Roman',
        'georgia': 'Georgia',
        'serif': 'Times New Roman',
        
        # Monospace fonts
        'monaco': 'Roboto Mono',
        'menlo': 'Roboto Mono',
        'courier new': 'Courier New',
        'courier': 'Courier New',
        'monospace': 'Roboto Mono',
        
        # Default mappings
        'sans-serif': 'Inter',
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