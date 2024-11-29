import pandas as pd

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

def calculate_manual_handling(teeth):
    """Calculate the manual handling banner based on the weight."""
    weight = 1 if teeth == 'RMD' else teeth * 0.145
    if weight > 30:
        return ("Very Heavy Die - Use Lifting Aids", "(Over 30kg)")
    elif weight > 25:
        return ("Heavy Die - Use Lifting Aids", "(Over 25kg)")
    else:
        return ("Within Manual Handling Lifting Range", "(Male = up to 25kg, Female = up to 16kg)")
Full Updated Code with get_label_data and GRF Handling:
python
Copy code
import os
import pandas as pd
import re  # Regular expressions for input validation
import win32print  # Import for Windows printer handling

def load_grf_to_zpl(filename):
    """Load GRF data from the file and prepare it for sending to the printer."""
    with open(filename, 'r') as file:
        grf_data = file.read().strip()

    # Parse metadata from the GRF file (if needed)
    total_bytes = len(grf_data) // 2  # Each byte is represented by 2 hex characters
    bytes_per_row = 64  # Modify this based on your specific image format (8 pixels per byte)
    
    # Compose the ~DG command to store the image in the printer
    image_name = os.path.basename(filename).split('.')[0].upper()  # Use filename as image name
    return f"~DG{image_name}.GRF,{total_bytes},{bytes_per_row},{grf_data}"

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

def calculate_manual_handling(teeth):
    """Calculate the manual handling banner based on the weight."""
    weight = 1 if teeth == 'RMD' else teeth * 0.145
    if weight > 30:
        return ("Very Heavy Die - Use Lifting Aids", "(Over 30kg)")
    elif weight > 25:
        return ("Heavy Die - Use Lifting Aids", "(Over 25kg)")
    else:
        return ("Within Manual Handling Lifting Range", "(Male = up to 25kg, Female = up to 16kg)")

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

    # Use GRF images for weight classifications and load weight images
    weight = label_data['weight']

    # Only load GRF images if conditions are met
    if weight > 30:
        zpl_commands.append(load_grf_to_zpl("1Weight.grf"))  # Load GRF for Very Heavy
        zpl_commands.append("^FO100,100^XG1WEIGHT.GRF^FS")  # Print the 1Weight image
    elif weight > 25:
        zpl_commands.append(load_grf_to_zpl("2Weight.grf"))  # Load GRF for Heavy
        zpl_commands.append("^FO100,100^XG2WEIGHT.GRF^FS")  # Print the 2Weight image
    else:
        zpl_commands.append(load_grf_to_zpl("3Weight.grf"))  # Load GRF for Manual Handling
        zpl_commands.append("^FO100,100^XG3WEIGHT.GRF^FS")  # Print the 3Weight image

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
        
        # Send to the printer
        # send_to_printer(zpl_output, printer_name)

if __name__ == "__main__":
    main()
