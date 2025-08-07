"""
Figma Data Processor
Handles incoming Figma plugin data and prepares it for PPTX generation
"""

def process_figma_data(figma_data):
    """
    Process the raw Figma data from the plugin
    
    Args:
        figma_data: Dictionary containing slides and metadata from Figma plugin
        
    Returns:
        Processed data ready for PPTX generation
    """
    processed_slides = []
    
    for slide in figma_data.get('slides', []):
        processed_slide = {
            'id': slide.get('id'),
            'name': slide.get('name', 'Untitled Slide'),
            'width': slide.get('width', 1920),
            'height': slide.get('height', 1080),
            'background': slide.get('background'),
            'layers': process_layers(slide.get('layers', [])),
            'frameInfo': slide.get('frameInfo', {}),
            'metadata': slide.get('metadata', {})
        }
        processed_slides.append(processed_slide)
    
    return {
        'fileName': figma_data.get('fileName', 'Untitled Presentation'),
        'slides': processed_slides,
        'config': {
            'slideWidth': 10,      # inches
            'slideHeight': 5.625,  # 16:9 aspect ratio
            'safeMargin': 0.3      # inches
        }
    }

def process_layers(layers):
    """
    Process individual layers from Figma
    
    Args:
        layers: List of layer objects from Figma
        
    Returns:
        Processed layers with enhanced metadata
    """
    processed_layers = []
    
    for layer in layers:
        processed_layer = {
            'type': layer.get('type'),
            'id': layer.get('id'),
            'name': layer.get('name', 'Untitled Layer'),
            'position': layer.get('position', {}),
            'relativePosition': layer.get('relativePosition', {}),
            'zIndex': layer.get('zIndex', 0),
            'depth': layer.get('depth', 0),
            'style': layer.get('style', {}),
            'content': layer.get('content', ''),
            'imageData': layer.get('imageData'),
            'shapeType': layer.get('shapeType')
        }
        
        # Validate layer data
        if processed_layer['type'] in ['TEXT', 'SHAPE', 'IMAGE']:
            processed_layers.append(processed_layer)
        else:
            print(f"Warning: Skipping unknown layer type: {processed_layer['type']}")
    
    # Sort layers by z-index (bottom to top)
    processed_layers.sort(key=lambda x: x.get('zIndex', 0))
    
    return processed_layers

def validate_figma_data(figma_data):
    """
    Validate the incoming Figma data
    
    Args:
        figma_data: Raw data from Figma plugin
        
    Returns:
        Boolean indicating if data is valid
    """
    if not isinstance(figma_data, dict):
        return False
    
    if not figma_data.get('slides'):
        return False
    
    for slide in figma_data['slides']:
        if not slide.get('layers'):
            continue
            
        for layer in slide['layers']:
            if not layer.get('type') or not layer.get('position'):
                return False
    
    return True 