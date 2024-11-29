import os  # Provides functions for interacting with the operating system, such as file handling
import pandas as pd  # Library for data manipulation and analysis, especially for handling data in tabular format (like Excel)
import re  # Regular expression library for string matching and manipulation
import win32print  # Windows-specific library for printing to printers
import tkinter as tk  # Standard GUI toolkit for creating desktop applications in Python
from tkinter import messagebox, filedialog  # Specific modules from tkinter for displaying messages and file dialogs


# Function to load ZPL commands from a specified file
def load_zpl_from_file(filename):
    with open(filename, 'r') as file:  # Open the file in read mode
        return file.read().strip()  # Read contents and remove any surrounding whitespace

# Function to retrieve label data based on a given ToolCode
def get_label_data(tool_code):
    df = pd.read_excel("RadiusDieAlternativesReport.xlsx")  # Load data from the Excel file
    row = df[df['ToolCode'] == tool_code]  # Filter row matching the provided ToolCode

    if row.empty:  # If no matching row is found
        return None  # Return None

    # Extract individual values from the DataFrame
    teeth = row['Teeth'].values[0]  # Number of teeth
    press_type = row['PressType'].values[0] if 'PressType' in row else 'Unknown'  # Get press type or default to 'Unknown'
    tool_type_code = row['ToolTypeCode'].values[0]  # Get tool type code
    plant_list = row['PlantList'].values[0]  # Get plant list
    
    # Calculate weight and manual handling guidelines
    weight = calculate_weight(teeth, press_type, tool_type_code)
    manual_handling = calculate_manual_handling(teeth, press_type, tool_type_code)

    # Return a dictionary containing all relevant data
    return {
        'teeth': teeth,
        'manual_handling': manual_handling,
        'spec_code': row['SpecCode'].values[0],
        'tool_code': tool_code,
        'legacy_code': row['LegacyCode'].values[0],
        'tool_code_text': tool_code,
        'press_type': press_type,
        'weight': weight,
        'plant_list': plant_list
    }

# Function to calculate weight based on the number of teeth, press type, and tool type code
def calculate_weight(teeth, press_type, tool_type_code):
    if tool_type_code == 'RMD':  # Special case for RMD tool type
        return 1  # Return a weight of 1

    # Define weight factors based on press size
    weight_factors = {
        "7 Inch": 0.145,
        "10 Inch": 0.232,
        "14 Inch": 0.305,
        "17 Inch": 0.378,
    }

    # Ensure that press_type is treated as a string
    press_type = str(press_type)

    # Calculate and return the total weight
    return teeth * weight_factors.get(press_type, 0)

# Function to calculate manual handling guidelines based on weight
def calculate_manual_handling(teeth, press_type, tool_type_code):
    weight = calculate_weight(teeth, press_type, tool_type_code)  # Calculate the weight

    # Determine handling recommendations based on weight
    if weight > 30:
        return ("Very Heavy Die - Use Lifting Aids", "(Over 30kg)")  # Suggest lifting aids for very heavy dies
    elif weight > 25:
        return ("Heavy Die - Use Lifting Aids", "(Over 25kg)")  # Suggest lifting aids for heavy dies
    else:
        return ("Within Manual Handling Lifting Range", "(Male = up to 25kg, Female = up to 16kg)")  # Safe handling range

# Function to generate ZPL commands for the label
def generate_zpl(label_data):
    zpl_commands = ["^XA"]  # Start ZPL commands

    # Add RFID command for the Preston plant
    if label_data['plant_list'] == 'Preston':
        rfid_code = f"{label_data['tool_code']}*00000"  # Construct RFID code
        zpl_commands.append(f"^RS8^RFW,A^FD{rfid_code}^FS^RFR,A")  # Add RFID command to ZPL

    # Determine which weight graphic to use based on calculated weight
    weight = label_data['weight']
    if weight > 30:
        zpl_commands.append(load_zpl_from_file("3Weight.grf"))  # Load graphic for very heavy weight
    elif weight > 25:
        zpl_commands.append(load_zpl_from_file("2Weight.grf"))  # Load graphic for heavy weight
    else:
        zpl_commands.append(load_zpl_from_file("1Weight.grf"))  # Load graphic for standard weight

    # Add weight and teeth count to ZPL
    zpl_commands.append(f"^FO599.703,37.963^A0N,39,39^FD{weight:.2f} KG^FS")  # Display weight in KG
    zpl_commands.append(f"^FO295.853,37.816^A0N,102,102^FD{label_data['teeth']} T^FS")  # Display teeth count

    # Calculate X positions for manual handling text
    line1_x_pos = calculate_x_position(label_data['manual_handling'][0], font_size=34, label_width_mm=100)
    line2_x_pos = calculate_x_position(label_data['manual_handling'][1], font_size=34, label_width_mm=100)

    # Add manual handling text to ZPL
    zpl_commands.append(f"^FO{line1_x_pos},163.119^A0N,34,34^FD{label_data['manual_handling'][0]}^FS")
    zpl_commands.append(f"^FO{line2_x_pos},200.119^A0N,34,34^FD{label_data['manual_handling'][1]}^FS")
    zpl_commands.append(f"^FO239.290,255.621^A0N,34,34^FD{label_data['spec_code']}^FS")
    zpl_commands.append(f"^FO239.811,315.209^BY2,2.0,80^BCN,N,N,Y,N,N^FD{label_data['tool_code']}^FS")  # Add barcode for ToolCode
    zpl_commands.append(f"^FO559.848,255.622^A0N,39,39^FDLegacy Code:^FS")  # Label for Legacy Code
    zpl_commands.append(f"^FO559.526,335.728^A0N,39,39^FD{label_data['legacy_code']}^FS")  # Display Legacy Code
    zpl_commands.append(f"^FO126.435,427.019^A0N,102,102^FD{label_data['tool_code_text']}^FS")  # Display ToolCode text
    zpl_commands.append("^XZ")  # End ZPL commands
    
    return "\n".join(zpl_commands)  # Return complete ZPL commands as a single string

# Function to calculate the X position for centring text on the label
def calculate_x_position(text, font_size=34, label_width_mm=100):
    label_width = label_width_mm * 8  # Convert label width from mm to pixels
    text_length = len(text) * font_size * 0.5  # Calculate length of text in pixels
    x_position = (label_width - text_length) / 2  # Centre text by calculating X position
    return max(int(x_position), 0)  # Return X position, ensuring it's not negative

# Function to save ZPL output to a text file for debugging
def save_zpl_to_file(zpl_output, tool_code):
    filename = f"zpl_output_{tool_code}.txt"  # Construct filename based on ToolCode
    with open(filename, 'w') as file:  # Open the file in write mode
        file.write(zpl_output)  # Write ZPL output to file
    print(f"ZPL output saved to {filename} for debugging.")  # Inform user of saved file

# Function to send ZPL output to the specified printer
def send_to_printer(zpl_output, printer_name, tool_code):
    try:
        # Save ZPL output to a file for debugging
        save_zpl_to_file(zpl_output, tool_code)
        
        printer_handle = win32print.OpenPrinter(printer_name)  # Open the printer
        try:
            job = win32print.StartDocPrinter(printer_handle, 1, ("ZPL Label", None, "RAW"))  # Start print job
            win32print.StartPagePrinter(printer_handle)  # Start new page
            win32print.WritePrinter(printer_handle, zpl_output.encode('utf-8'))  # Write ZPL output to printer
            win32print.EndPagePrinter(printer_handle)  # End page
            win32print.EndDocPrinter(printer_handle)  # End document
        finally:
            win32print.ClosePrinter(printer_handle)  # Close printer handle
        messagebox.showinfo("Success", "Label sent to printer.")  # Show success message
    except Exception as e:
        messagebox.showerror("Error", f"Error sending to printer: {e}")  # Show error message if printing fails

# Class to create the GUI for the label printer application using Tkinter
class LabelPrinterApp:
    def __init__(self, root):
        self.root = root  # Store reference to the main window
        self.root.title("Label Printer App")  # Set the title of the window

        # Printer selection label and dropdown
        tk.Label(root, text="Select Printer:").grid(row=0, column=0, padx=10, pady=5)  # Label for printer selection
        self.printer_name = tk.StringVar()  # Variable to hold selected printer name
        # Get a list of available printers
        printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        self.printer_dropdown = tk.OptionMenu(root, self.printer_name, *printers)  # Dropdown for printer selection
        self.printer_dropdown.grid(row=0, column=1, padx=10, pady=5)  # Position dropdown in the grid

        # ToolCode entry field
        tk.Label(root, text="Enter ToolCodes (comma-separated):").grid(row=1, column=0, padx=10, pady=5)  # Label for ToolCode entry
        self.tool_code_entry = tk.Entry(root, width=40)  # Entry field for ToolCodes
        self.tool_code_entry.grid(row=1, column=1, padx=10, pady=5)  # Position entry field in the grid

        # Print button
        self.print_button = tk.Button(root, text="Print Label", command=self.print_label)  # Button to print label
        self.print_button.grid(row=2, column=0, columnspan=2, pady=10)  # Position button in the grid

    # Function to handle the printing of labels
    def print_label(self):
        printer_name = self.printer_name.get()  # Get selected printer name
        tool_codes = [code.strip() for code in self.tool_code_entry.get().split(",")]  # Split ToolCodes and remove extra spaces

        if not printer_name:  # Check if a printer is selected
            messagebox.showerror("Error", "Please select a printer.")  # Show error if no printer is selected
            return

        # Process each ToolCode entered by the user
        for tool_code in tool_codes:
            if not validate_tool_code(tool_code):  # Validate the ToolCode format
                messagebox.showwarning("Warning", f"Invalid ToolCode: {tool_code}")  # Show warning for invalid ToolCode
                continue  # Skip to the next ToolCode

            label_data = get_label_data(tool_code)  # Get label data for the ToolCode
            if label_data:  # If label data is found
                zpl_output = generate_zpl(label_data)  # Generate ZPL commands
                send_to_printer(zpl_output, printer_name, label_data['tool_code'])  # Send ZPL to printer
            else:
                messagebox.showwarning("Warning", f"ToolCode {tool_code} not found.")  # Show warning if ToolCode not found

# Function to validate the ToolCode format
def validate_tool_code(tool_code):
    if len(tool_code) > 10 or not re.match("^[a-zA-Z0-9]*$", tool_code):  # Check length and valid characters
        return False  # Return False if invalid
    return True  # Return True if valid

# Main entry point for running the application
if __name__ == "__main__":
    root = tk.Tk()  # Create the main window
    app = LabelPrinterApp(root)  # Instantiate the LabelPrinterApp
    root.mainloop()  # Start the GUI event loop
