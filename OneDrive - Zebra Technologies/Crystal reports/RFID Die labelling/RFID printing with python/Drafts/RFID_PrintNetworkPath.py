import os 
import pandas as pd
import re  # Regular expressions for input validation
import win32print  # Import for Windows printer handling

def load_hex_from_file(filename):
    """Load hex data and image metadata from the file and format it for ZPL."""
    with open(filename, 'r') as file:
        # Read the first line for metadata (width, height, mode)
        metadata = file.readline().strip()
        width, height, mode = metadata.split(',')
        width, height = int(width), int(height)
        
        # Read the remaining hex data
        hex_data = file.read().strip()  # Read and strip whitespace

    # Calculate width in bytes for ZPL
    # One byte is 8 pixels, and mode can influence byte width
    bytes_per_row = width * (4 if mode == 'RGBA' else 3)  # Assuming 4 bytes per pixel for RGBA, 3 for RGB
    total_bytes = len(hex_data) // 2  # Divide by 2 since hex is two characters per byte

    # Return formatted ZPL command with hex data
    return f"^GFA,{total_bytes},{total_bytes},{bytes_per_row},{hex_data}"

def calculate_manual_handling(teeth):
    """Calculate the manual handling banner based on the weight."""
    weight = 1 if teeth == 'RMD' else teeth * 0.145
    if weight > 30:
        return ("Very Heavy Die - Use Lifting Aids", "(Over 30kg)")
    elif weight > 25:
        return ("Heavy Die - Use Lifting Aids", "(Over 25kg)")
    else:
        return ("Within Manual Handling Lifting Range", "(Male = up to 25kg, Female = up to 16kg)")

def get_label_data(tool_code):
    """Load label data from the Excel spreadsheet based on ToolCode."""
    df = pd.read_excel("RadiusDieAlternativesReport.xlsx")  # Update with your actual file path
    row = df[df['ToolCode'] == tool_code]

    if row.empty:
        print(f"ToolCode {tool_code} not found.")
        return None

    teeth = row['Teeth'].values[0]
    return {
        'weight': 1 if row['ToolTypeCode'].values[0] == 'RMD' else teeth * 0.145,
        'teeth': teeth,
        'manual_handling': calculate_manual_handling(teeth),
        'spec_code': row['SpecCode'].values[0],
        'tool_code': tool_code,
        'legacy_code': row['LegacyCode'].values[0],
        'tool_code_text': tool_code  # Assuming ToolCode is the same as the input ToolCode
    }

def generate_zpl(label_data):
    """Generate ZPL commands based on the label data."""
    zpl_commands = []
    zpl_commands.append("^XA")  # Start the ZPL command

    # RFID Encryption
    rfid_code = f"{label_data['tool_code']}*00000"
    zpl_commands.append("^RS8")
    zpl_commands.append("^RFW,A")
    zpl_commands.append(f"^FD{rfid_code}^FS")
    zpl_commands.append("^RFR,A")

    # Use text for weight classifications and load weight images
    weight = label_data['weight']

    # Only load images if conditions are met
    if weight > 30:
        zpl_commands.append(load_hex_from_file("1Weight_hex.txt"))  # Load image for Very Heavy
    elif weight > 25:
        zpl_commands.append(load_hex_from_file("2Weight_hex.txt"))  # Load image for Heavy
    else:
        zpl_commands.append(load_hex_from_file("3Weight_hex.txt"))  # Load image for Manual Handling

    # Update ZPL commands with new sizes and positions based on the updated chart
    zpl_commands.append(f"^FO666.703,37.963^A0N,34,34^FD{weight:.2f} KG^FS")  # Weight
    zpl_commands.append(f"^FO311.853,64.816^A0N,51,51^FD{label_data['teeth']} T^FS")  # Teeth

    # Get the manual handling text lines - Only add it once
    manual_handling_lines = label_data['manual_handling']
    zpl_commands.append(f"^FO147.934,163.119^A0N,34,34^FD{manual_handling_lines[0]}^FS")  # Line 1
    zpl_commands.append(f"^FO147.934,200.119^A0N,34,34^FD{manual_handling_lines[1]}^FS")  # Line 2

    # Other label data
    zpl_commands.append(f"^FO159.290,255.621^A0N,34,34^FD{label_data['spec_code']}^FS")  # SpecCode
    zpl_commands.append(f"^FO151.811,315.209^BY2,2.0,80^BCN,N,N,Y,N,N^FD{label_data['tool_code']}^FS")  # ToolCode Barcode
    zpl_commands.append(f"^FO639.848,239.622^A0N,31,31^FDLegacy Code:^FS")  # Legacy Code Text
    zpl_commands.append(f"^FO639.526,335.728^A0N,39,39^FD{label_data['legacy_code']}^FS")  # Legacy Code
    zpl_commands.append(f"^FO126.435,427.019^A0N,102,102^FD{label_data['tool_code_text']}^FS")  # ToolCode

    # End the label formatting
    zpl_commands.append("^XZ")  # End of the label

    return "\n".join(zpl_commands)

def validate_tool_code(tool_code):
    """Validate the ToolCode to ensure it meets criteria."""
    if len(tool_code) > 10:
        print("Error: ToolCode must be 10 characters or less.")
        return False
    if not re.match("^[a-zA-Z0-9]*$", tool_code):
        print("Error: ToolCode can only contain letters and numbers (no special characters).")
        return False
    return True

def send_to_printer(zpl_output, printer_name):
    """Send ZPL output to the network printer."""
    try:
        printer_handle = win32print.OpenPrinter(printer_name)
        try:
            job = win32print.StartDocPrinter(printer_handle, 1, ("ZPL Label", None, "RAW"))
            win32print.StartPagePrinter(printer_handle)
            win32print.WritePrinter(printer_handle, zpl_output.encode('utf-8'))
            win32print.EndPagePrinter(printer_handle)
            win32print.EndDocPrinter(printer_handle)
        finally:
            win32print.ClosePrinter(printer_handle)
    except Exception as e:
        print(f"An error occurred while sending to the printer: {e}")

def main():
    # Input Tool Code with validation
    while True:
        tool_code_input = input("Enter ToolCode (max 10 characters, letters and numbers only): ")
        if validate_tool_code(tool_code_input):
            break  # Exit loop if validation is successful

    label_data = get_label_data(tool_code_input)

    if label_data:
        # Generate ZPL commands
        zpl_output = generate_zpl(label_data)

        # Display the ZPL output on the console
        print("\nGenerated ZPL Commands:\n")
        print(zpl_output)  # Display the generated ZPL commands

        # Write the ZPL output to zpl_output.txt (overwriting it)
        with open("zpl_output.txt", 'w') as file:  # Open in write mode to overwrite
            file.write(zpl_output)

        # Define the network printer's name or path (for future use)
        printer_name = r"\\09sp-prntinf01\99j213003422"  # Update with your actual printer path or name 
        
        # Comment out the actual sending to the printer for now
        send_to_printer(zpl_output, printer_name)

if __name__ == "__main__":
    main()
