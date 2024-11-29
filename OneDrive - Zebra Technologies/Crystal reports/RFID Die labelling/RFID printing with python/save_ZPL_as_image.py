import requests

def render_zpl_with_labelary(zpl_command):
    """Send ZPL command to Labelary API and save the returned image."""
    url = "http://api.labelary.com/v1/printer/0/label"  # Correct URL for Labelary API
    headers = {'Accept': 'image/png'}

    # Send the ZPL command
    response = requests.post(url, headers=headers, data=zpl_command)

    if response.status_code == 200:
        # Save the image
        with open("label_image.png", "wb") as img_file:
            img_file.write(response.content)
        print("Label image saved as 'label_image.png'")
    else:
        # Print the error details
        print("Error rendering ZPL:", response.status_code, response.text)

def main():
    # Read ZPL command from the text file
    try:
        with open("zpl_output.txt", "r") as file:
            zpl_command = file.read()
    except FileNotFoundError:
        print("Error: The file 'zpl_output.txt' was not found.")
        return

    # Render ZPL command using Labelary API
    render_zpl_with_labelary(zpl_command)

if __name__ == "__main__":
    main()
