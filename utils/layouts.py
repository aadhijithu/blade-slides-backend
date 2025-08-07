"""
Layout and positioning utilities for precise Figma to PowerPoint conversion
"""

def calculate_positioning(slide_data, config):
    """
    Calculate positioning parameters for a slide
    
    Args:
        slide_data: Slide data from Figma
        config: Configuration with slide dimensions
        
    Returns:
        Dictionary with positioning calculations
    """
    # Get slide dimensions
    figma_width = slide_data.get('width', 1920)
    figma_height = slide_data.get('height', 1080)
    
    # PowerPoint slide dimensions (in inches)
    ppt_width = config['slideWidth']
    ppt_height = config['slideHeight']
    safe_margin = config['safeMargin']
    
    # Available area (excluding margins)
    available_width = ppt_width - (2 * safe_margin)
    available_height = ppt_height - (2 * safe_margin)
    
    # Calculate aspect ratios
    figma_aspect = figma_width / figma_height
    ppt_aspect = available_width / available_height
    
    # Calculate scaling and offsets for aspect ratio preservation
    if figma_aspect > ppt_aspect:
        # Figma frame is wider - fit to width
        scale = available_width / figma_width
        content_height = figma_height * scale
        offset_x = safe_margin
        offset_y = safe_margin + (available_height - content_height) / 2
    else:
        # Figma frame is taller - fit to height
        scale = available_height / figma_height
        content_width = figma_width * scale
        offset_x = safe_margin + (available_width - content_width) / 2
        offset_y = safe_margin
    
    return {
        'scale': scale,
        'offset_x': offset_x,
        'offset_y': offset_y,
        'figma_width': figma_width,
        'figma_height': figma_height,
        'ppt_width': ppt_width,
        'ppt_height': ppt_height,
        'available_width': available_width,
        'available_height': available_height
    }

def calculate_layer_position(layer, positioning, config):
    """
    Calculate precise position for a layer in PowerPoint coordinates
    
    Args:
        layer: Layer data from Figma
        positioning: Positioning calculations from calculate_positioning
        config: Configuration settings
        
    Returns:
        Dictionary with x, y, width, height in inches
    """
    # Try relative positioning first (more accurate)
    if layer.get('relativePosition'):
        rel_pos = layer['relativePosition']
        
        # Convert relative position (0-1) to PowerPoint inches
        x = positioning['offset_x'] + (rel_pos['x'] * positioning['available_width'])
        y = positioning['offset_y'] + (rel_pos['y'] * positioning['available_height'])
        width = rel_pos['width'] * positioning['available_width']
        height = rel_pos['height'] * positioning['available_height']
    
    # Fallback to absolute positioning with scaling
    elif layer.get('position'):
        pos = layer['position']
        scale = positioning['scale']
        
        x = positioning['offset_x'] + (pos['x'] * scale)
        y = positioning['offset_y'] + (pos['y'] * scale)
        width = pos['width'] * scale
        height = pos['height'] * scale
    
    else:
        # No position data - place at origin with minimum size
        x = positioning['offset_x']
        y = positioning['offset_y']
        width = 1.0
        height = 0.5
    
    # Ensure minimum sizes for visibility
    width = max(width, 0.1)
    height = max(height, 0.1)
    
    # Boundary checks to prevent elements going off-slide
    x = max(0, min(x, positioning['ppt_width'] - width))
    y = max(0, min(y, positioning['ppt_height'] - height))
    
    return {
        'x': x,
        'y': y,
        'width': width,
        'height': height
    }

def scale_font_size(figma_font_size, scale_factor):
    """
    Scale font size from Figma to PowerPoint
    
    Args:
        figma_font_size: Font size in Figma (pixels)
        scale_factor: Scaling factor from positioning calculations
        
    Returns:
        Scaled font size in points
    """
    # Convert pixels to points (approximate: 1pt = 1.33px)
    base_pt_size = figma_font_size * 0.75
    
    # Apply scaling
    scaled_size = base_pt_size * scale_factor
    
    # Ensure readable font size
    return max(8, min(72, scaled_size))

def pixels_to_inches(pixels, dpi=96):
    """
    Convert pixels to inches
    
    Args:
        pixels: Size in pixels
        dpi: Dots per inch (default 96 for screen)
        
    Returns:
        Size in inches
    """
    return pixels / dpi

def inches_to_pixels(inches, dpi=96):
    """
    Convert inches to pixels
    
    Args:
        inches: Size in inches
        dpi: Dots per inch (default 96 for screen)
        
    Returns:
        Size in pixels
    """
    return inches * dpi

def calculate_text_box_size(text_content, font_size, font_family='Arial'):
    """
    Estimate text box size based on content
    
    Args:
        text_content: Text content
        font_size: Font size in points
        font_family: Font family name
        
    Returns:
        Dictionary with estimated width and height in inches
    """
    if not text_content:
        return {'width': 1.0, 'height': 0.3}
    
    # Rough estimation based on character count and font size
    char_width_ratio = 0.6  # Average character width ratio to font size
    line_height_ratio = 1.2  # Line height ratio to font size
    
    # Estimate width based on longest line
    lines = text_content.split('\n')
    max_line_length = max(len(line) for line in lines) if lines else 0
    
    estimated_width_pt = max_line_length * font_size * char_width_ratio
    estimated_height_pt = len(lines) * font_size * line_height_ratio
    
    # Convert points to inches (72 points = 1 inch)
    width_inches = max(0.5, estimated_width_pt / 72)
    height_inches = max(0.3, estimated_height_pt / 72)
    
    return {
        'width': width_inches,
        'height': height_inches
    }

def adjust_for_safe_area(x, y, width, height, config):
    """
    Adjust position to keep within safe area
    
    Args:
        x, y: Position in inches
        width, height: Size in inches
        config: Configuration with safe margins
        
    Returns:
        Adjusted position dictionary
    """
    safe_margin = config.get('safeMargin', 0.3)
    slide_width = config.get('slideWidth', 10)
    slide_height = config.get('slideHeight', 5.625)
    
    # Adjust position to keep within safe area
    min_x = safe_margin
    max_x = slide_width - safe_margin - width
    min_y = safe_margin
    max_y = slide_height - safe_margin - height
    
    adjusted_x = max(min_x, min(x, max_x))
    adjusted_y = max(min_y, min(y, max_y))
    
    return {
        'x': adjusted_x,
        'y': adjusted_y,
        'width': width,
        'height': height
    }

def get_slide_aspect_ratio(config):
    """
    Get slide aspect ratio
    
    Args:
        config: Configuration dictionary
        
    Returns:
        Aspect ratio as float
    """
    width = config.get('slideWidth', 10)
    height = config.get('slideHeight', 5.625)
    return width / height

def is_landscape_orientation(width, height):
    """
    Check if dimensions represent landscape orientation
    
    Args:
        width, height: Dimensions
        
    Returns:
        Boolean indicating landscape orientation
    """
    return width > height 