from PIL import Image, ImageDraw, ImageFont
import os

def create_icon(size, text, bg_color, fg_color):
    """Create a single icon with text"""
    try:
        font = ImageFont.truetype("arial.ttf", int(size * 0.6))
    except IOError:
        font = ImageFont.load_default()

    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Draw a rounded rectangle
    draw.rounded_rectangle([2, 2, size-2, size-2], radius=size//4, fill=bg_color)
    
    # Calculate text position
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = (size - text_width) / 2
    y = (size - text_height) / 2
    
    # Draw text
    draw.text((x, y), text, fill=fg_color, font=font)
    return img

def create_ico_file():
    """Create and save icons as a .ico file"""
    # Define icon sizes (Windows typically uses these sizes)
    sizes = [16, 32, 48, 64, 128, 256]
    
    # Create icons for each size
    icons = {
        'home': [],
        'about': [],
        'insert': []
    }
    
    # Colors
    bg_color = "#e94560"  # Red color from your app
    fg_color = "white"
    
    # Create icons for each size
    for size in sizes:
        # Home icon (H)
        home_icon = create_icon(size, "H", bg_color, fg_color)
        icons['home'].append(home_icon)
        
        # About icon (i)
        about_icon = create_icon(size, "i", bg_color, fg_color)
        icons['about'].append(about_icon)
        
        # Insert file icon (F)
        insert_icon = create_icon(size, "F", bg_color, fg_color)
        icons['insert'].append(insert_icon)
    
    # Save each icon set as a separate .ico file
    icons['home'][0].save('home.ico', format='ICO', sizes=[(s, s) for s in sizes])
    icons['about'][0].save('about.ico', format='ICO', sizes=[(s, s) for s in sizes])
    icons['insert'][0].save('insert.ico', format='ICO', sizes=[(s, s) for s in sizes])
    
    print("Icon files created successfully!")

if __name__ == "__main__":
    create_ico_file() 