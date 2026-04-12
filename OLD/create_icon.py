"""
Create an icon for the NetSuite Import Processor
"""
from PIL import Image, ImageDraw, ImageFont
import os

def create_icon():
    # Create multiple sizes for ICO file
    sizes = [16, 32, 48, 64, 128, 256]
    images = []
    
    for size in sizes:
        # Create image with gradient background
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        # Draw rounded rectangle background with blue gradient effect
        # Main background - dark blue
        draw.rounded_rectangle(
            [(1, 1), (size-2, size-2)],
            radius=size//6,
            fill=(30, 60, 114, 255)  # Dark blue
        )
        
        # Inner lighter area
        margin = size // 8
        draw.rounded_rectangle(
            [(margin, margin), (size-margin-1, size-margin-1)],
            radius=size//8,
            fill=(45, 85, 150, 255)  # Medium blue
        )
        
        # Draw "NS" text
        try:
            # Try to use a nice font
            font_size = size // 2
            try:
                font = ImageFont.truetype("arial.ttf", font_size)
            except:
                try:
                    font = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", font_size)
                except:
                    font = ImageFont.load_default()
        except:
            font = ImageFont.load_default()
        
        text = "NS"
        
        # Get text bounding box
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        # Center the text
        x = (size - text_width) // 2
        y = (size - text_height) // 2 - bbox[1]
        
        # Draw text with shadow
        shadow_offset = max(1, size // 32)
        draw.text((x + shadow_offset, y + shadow_offset), text, fill=(20, 40, 80, 200), font=font)
        draw.text((x, y), text, fill=(255, 255, 255, 255), font=font)
        
        images.append(img)
    
    # Save as ICO
    icon_path = os.path.join(os.path.dirname(__file__), 'icon.ico')
    images[0].save(
        icon_path,
        format='ICO',
        sizes=[(s, s) for s in sizes],
        append_images=images[1:]
    )
    print(f"Icon created: {icon_path}")
    return icon_path

if __name__ == "__main__":
    create_icon()

