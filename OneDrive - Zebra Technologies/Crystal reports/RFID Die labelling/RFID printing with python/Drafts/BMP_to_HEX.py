from PIL import Image

def png_to_hex(png_file, hex_file):
    """Convert a PNG image to its hexadecimal representation and save to a text file."""
    # Load the image using Pillow
    img = Image.open(png_file)
    
    # Get the image mode and size
    width, height = img.size
    mode = img.mode
    
    # Convert the image to raw data
    raw_data = img.tobytes()
    
    # Convert raw data to hexadecimal
    hex_data = raw_data.hex()
    
    # Write image metadata (width, height, mode) and hex data to the file
    with open(hex_file, 'w') as hex_out:
        hex_out.write(f"{width},{height},{mode}\n")
        hex_out.write(hex_data)
    
    print(f"Hexadecimal data with metadata saved to {hex_file}.")

def hex_to_png(hex_file, output_png_file):
    """Convert a hexadecimal representation back to a PNG image."""
    with open(hex_file, 'r') as hex_in:
        # Read the first line for metadata (width, height, mode)
        metadata = hex_in.readline().strip()
        width, height, mode = metadata.split(',')
        width, height = int(width), int(height)
        
        # Read the hex data
        hex_data = hex_in.read().strip()
    
    # Convert the hex string back to raw bytes
    raw_data = bytes.fromhex(hex_data)
    
    # Create a new image from the raw bytes using the stored metadata
    img = Image.frombytes(mode, (width, height), raw_data)
    
    # Save the image as a PNG
    img.save(output_png_file)
    
    print(f"Image created from hex data and saved as {output_png_file}.")

def main():
    # Specify the image files
    images = [
        ('1Weight.png', '1Weight_hex.txt'),
        ('2Weight.png', '2Weight_hex.txt'),
        ('3Weight.png', '3Weight_hex.txt')
    ]
    
    # Convert PNG images to hex files
    for png_file, hex_file in images:
        png_to_hex(png_file, hex_file)

    # Convert hex files back to PNG images
    hex_files = [
        ('1Weight_hex.txt', '1Weight_converted.png'),
        ('2Weight_hex.txt', '2Weight_converted.png'),
        ('3Weight_hex.txt', '3Weight_converted.png')
    ]

    for hex_file, output_png_file in hex_files:
        hex_to_png(hex_file, output_png_file)

if __name__ == "__main__":
    main()
