from PIL import Image, ImageDraw, ImageFont
import os

def create_icon():
    """Create icon for MontanaSpecimensMapper application"""
    # Create a new image with a white background
    size = (256, 256)
    icon = Image.new('RGBA', size, (255, 255, 255, 0))
    draw = ImageDraw.Draw(icon)
    
    # Define colors
    montana_blue = "#2c3e50"  # Dark blue for main shape
    accent_color = "#3498db"  # Lighter blue for accent
    
    # Calculate dimensions for Montana silhouette
    padding = 20
    montana_width = size[0] - (2 * padding)
    montana_height = size[1] - (2 * padding)
    
    # Draw a simplified Montana shape (stylized rectangle with mountains)
    # Main shape - slightly tilted rectangle
    points = [
        (padding + 20, padding + 40),  # Top left
        (size[0] - padding - 20, padding + 20),  # Top right
        (size[0] - padding - 30, size[1] - padding - 20),  # Bottom right
        (padding + 10, size[1] - padding - 40)  # Bottom left
    ]
    draw.polygon(points, fill=montana_blue)
    
    # Add grid lines to represent counties
    grid_color = "#ffffff"
    line_width = 2
    
    # Horizontal grid lines
    for y in range(4):
        y_pos = padding + 40 + (y * ((size[1] - 2 * padding - 60) // 3))
        draw.line(
            [(padding + 15, y_pos), (size[0] - padding - 25, y_pos - 5)],
            fill=grid_color,
            width=line_width
        )
    
    # Vertical grid lines
    for x in range(4):
        x_pos = padding + 20 + (x * ((size[0] - 2 * padding - 40) // 3))
        draw.line(
            [(x_pos, padding + 30), (x_pos, size[1] - padding - 30)],
            fill=grid_color,
            width=line_width
        )
    
    # Add a point marker to represent data points
    marker_size = 20
    marker_pos = (size[0] // 2, size[1] // 2)
    draw.ellipse(
        [
            (marker_pos[0] - marker_size//2, marker_pos[1] - marker_size//2),
            (marker_pos[0] + marker_size//2, marker_pos[1] + marker_size//2)
        ],
        fill=accent_color
    )
    
    # Add a subtle shadow
    shadow = Image.new('RGBA', size, (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_offset = 8
    shadow_points = [
        (p[0] + shadow_offset, p[1] + shadow_offset) for p in points
    ]
    shadow_draw.polygon(shadow_points, fill=(0, 0, 0, 50))
    
    # Merge shadow and main icon
    icon = Image.alpha_composite(shadow, icon)
    
    # Save in different sizes
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    
    # Save as PNG
    icon.save('app_icon.png', 'PNG')
    
    # Save as ICO (Windows icon) with multiple sizes
    icon_sizes = []
    for size in sizes:
        resized = icon.resize(size, Image.Resampling.LANCZOS)
        icon_sizes.append(resized)
    
    icon_sizes[0].save('app_icon.ico', format='ICO', sizes=[(s[0], s[1]) for s in sizes])
    
    print("✓ Icon files created successfully:")
    print("  - app_icon.png")
    print("  - app_icon.ico")

if __name__ == "__main__":
    create_icon() 