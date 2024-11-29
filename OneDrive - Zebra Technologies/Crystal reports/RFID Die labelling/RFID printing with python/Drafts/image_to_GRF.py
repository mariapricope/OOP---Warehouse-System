from PIL import Image
import numpy as np
import os

# Function to round dimensions up to the nearest multiple of 8
def round_up_to_8(x):
    return ((x + 7) // 8) * 8

# Function to convert an image to GRF format
def image_to_grf(image_path, output_grf_path, rotate=None, width=None, height=None):
    # Open the image and convert to monochrome (1-bit, black and white)
    image = Image.open(image_path).convert("1")

    # Resize image if width and/or height are specified
    if width or height:
        new_width = width if width else image.width
        new_height = height if height else image.height
        image = image.resize((new_width, new_height))

    # Round image width to nearest multiple of 8 (required for ZPL)
    new_width = round_up_to_8(image.width)
    new_height = image.height

    # Resize the image to match the required width (multiple of 8)
    image = image.resize((new_width, new_height))

    # Rotate image if needed
    if rotate:
        image = image.rotate(rotate, expand=True)

    # Convert image to numpy array (1-bit black and white)
    image_array = np.array(image)

    # Flatten the 2D array into a 1D array and convert binary pixels to hexadecimal
    hex_data = ""
    for row in image_array:
        binary_row = ''.join(['1' if pixel == 0 else '0' for pixel in row])  # Invert pixel values for ZPL
        hex_row = format(int(binary_row, 2), f'0{new_width // 4}X')  # Convert to hex
        hex_data += hex_row

    # Calculate the total number of bytes in the image
    bytes_per_row = new_width // 8
    total_bytes = bytes_per_row * new_height

    # Generate the ZPL GRF command with the image data
    grf_data = f"~DGIMAGE.GRF,{total_bytes},{bytes_per_row},{hex_data}"

    # Save the GRF data to a file
    with open(output_grf_path, "w") as f:
        f.write(grf_data)

    print(f"GRF data saved to {output_grf_path}")

if __name__ == "__main__":
    # List of image files and corresponding output .grf files
    image_files = ["1Weight.png", "2Weight.png", "3Weight.png"]
    
    # Loop over each image, convert to GRF, and save
    for image_file in image_files:
        base_name = os.path.splitext(image_file)[0]  # Get the base name (without extension)
        output_grf_file = f"{base_name}.grf"  # Set the output file name to match the image
        
        image_to_grf(image_file, output_grf_file)

