import os
import pandas as pd
import re  # Regular expressions for input validation
import win32print  # Import for Windows printer handling

def load_zpl_from_file(filename):
    """Load ZPL commands from a text file."""
    with open(filename, 'r') as file:
        zpl_data = file.read().strip()
    return zpl_data

def get_label_data(tool_code):
    """Load label data from the Excel spreadsheet based on ToolCode."""
    df = pd.read_excel("RadiusDieAlternativesReport.xlsx")  # Update with your actual file path
    row = df[df['ToolCode'] == tool_code]

    if row.empty:
        print(f"ToolCode {tool_code} not found.")
        return None

    teeth = row['Teeth'].values[0]
    press_type = row['PressType'].values[0] if 'PressType' in df.columns else 'Unknown' #size of the press
    tool_type_code = row['ToolTypeCode'].values[0]   # RMD or RRD
    plant_list = row['PlantList'].values[0]

   

    # Calculate weight based on tool type and press size
 #   weight = 1 if tool_type_code == 'RMD' else (
 #       teeth * 0.145 if press_type == "7 Inch" else
 #       teeth * 0.232 if press_type == "10 Inch" else
 #       teeth * 0.305 if press_type == "14 Inch" else
 #       teeth * 0.378 if press_type == "17 Inch" else 0
#    )

    manual_handling = calculate_manual_handling(teeth, press_type, tool_type_code) # Pass all necessary parameters

    return {
    
        'teeth': teeth,
        'manual_handling': manual_handling,
        'spec_code': row['SpecCode'].values[0],
        'tool_code': tool_code,
        'legacy_code': row['LegacyCode'].values[0],
        'tool_code_text': tool_code,  # Assuming ToolCode is the same as the input ToolCode
        'press_type' : press_type,
        'weight' : weight,
        'plant_list' : plant_list # include PlantList in the returnd data
        
    }

def calculate_manual_handling(teeth, press_type, tool_type_code):

    """Calculate the manual handling banner based on the weight."""

    weight = 1 if tool_type_code == 'RMD' else (
        teeth * 0.145 if press_type == "7 Inch" else
        teeth * 0.232 if press_type == "10 Inch" else
        teeth * 0.305 if press_type == "14 Inch" else
        teeth * 0.378 if press_type == "17 Inch" else 0
    )


    
    if weight > 30:
        return ("Very Heavy Die - Use Lifting Aids", "(Over 30kg)")
    elif weight > 25:
        return ("Heavy Die - Use Lifting Aids", "(Over 25kg)")
    else:
        return ("Within Manual Handling Lifting Range", "(Male = up to 25kg, Female = up to 16kg)")

def calculate_x_position(text, font_size=34, label_width_mm=100):
    """Calculate the x-position to centre the text on the label (in mm)."""
    # Convert label width from mm to dots (assuming 8 dots per mm for Zebra printers)
    label_width = label_width_mm * 8  # 100mm -> 800 dots

    # Estimate text length more accurately (Zebra ZPL font widths can vary, this is an approximation)
    # The width of a single character is roughly 0.5 times the font size in dots for the default font.
    text_length = len(text) * font_size * 0.5  # 0.5 scaling factor for character width estimation
    x_position = (label_width - text_length) / 2  # Calculate the centred X position

    # Ensure the X position is not negative (if text is too long to fit)
    if x_position < 0:
        x_position = 0
    
    return int(x_position)

def generate_zpl(label_data):
    """Generate ZPL commands based on the label data."""
    zpl_commands = []
    zpl_commands.append("^XA")  # Start the ZPL command

    # RFID Encryption for Preston only:
    if label_data['plant_list'] == 'Preston':
        rfid_code = f"{label_data['tool_code']}*00000"
        zpl_commands.append("^RS8")
        zpl_commands.append("^RFW,A")
        zpl_commands.append(f"^FD{rfid_code}^FS")
        zpl_commands.append("^RFR,A")

    # Load ZPL commands for weight classifications
    weight = label_data['weight']

    if weight > 30:
        zpl_commands.append(load_zpl_from_file("3Weight.grf"))  # Very Heavy
    elif weight > 25:
        zpl_commands.append(load_zpl_from_file("2Weight.grf"))  # Heavy
    else:
        zpl_commands.append(load_zpl_from_file("1Weight.grf"))  # Manual Handling

    # Example weight and teeth
    zpl_commands.append(f"^FO599.703,37.963^A0N,39,39^FD{weight:.2f} KG^FS")  # Weight
    zpl_commands.append(f"^FO295.853,37.816^A0N,102,102^FD{label_data['teeth']} T^FS")  # Teeth

    # Get the manual handling text lines and calculate X positions for each line
    manual_handling_lines = label_data['manual_handling']
    

    # Calculate the X position for Line 1 and Line 2 to ensure both are centred
    line1_x_pos = calculate_x_position(manual_handling_lines[0], font_size=34, label_width_mm=100)
    line2_x_pos = calculate_x_position(manual_handling_lines[1], font_size=34, label_width_mm=100)


    # Add Line 1 and Line 2 with calculated centred X positions
    zpl_commands.append(f"^FO{line1_x_pos},163.119^A0N,34,34^FD{manual_handling_lines[0]}^FS")  # Line 1 (centred)
    zpl_commands.append(f"^FO{line2_x_pos},200.119^A0N,34,34^FD{manual_handling_lines[1]}^FS")  # Line 2 (centred)

    zpl_commands.append(f"^FO{line2_x_pos},200.119^A0N,34,34^FD{manual_handling_lines[1]}^FS")  # Line 2 (centred)

    # Other label details
    zpl_commands.append(f"^FO239.290,255.621^A0N,34,34^FD{label_data['spec_code']}^FS")  # SpecCode
    zpl_commands.append(f"^FO239.811,315.209^BY2,2.0,80^BCN,N,N,Y,N,N^FD{label_data['tool_code']}^FS")  # ToolCode Barcode
    zpl_commands.append(f"^FO559.848,255.622^A0N,39,39^FDLegacy Code:^FS")  # Legacy Code Text
    zpl_commands.append(f"^FO559.526,335.728^A0N,39,39^FD{label_data['legacy_code']}^FS")  # Legacy Code
    zpl_commands.append(f"^FO126.435,427.019^A0N,102,102^FD{label_data['tool_code_text']}^FS")  # ToolCode

    # End of label formatting
    zpl_commands.append("^XZ")  # End of label

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
        # Uncomment the following line to enable printing
        # send_to_printer(zpl_output, printer_name)

if __name__ == "__main__":
    main()
