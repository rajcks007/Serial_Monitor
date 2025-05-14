import os
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
import serial.tools.list_ports
from serial import Serial
import xml.etree.ElementTree as ET
import csv
import threading
import datetime
import db_loader
from QR_Scanner import barcode_scanner

# Function to get the path to the icon file, works for both script and EXE
def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS  # Path where PyInstaller extracts files
    else:
        base_path = os.path.dirname(__file__)  # Normal Python script path
    return os.path.join(base_path, relative_path)

# Function to set the icon for the Tkinter window
def set_window_icon(window, icon_filename):
    icon_path = resource_path(icon_filename)  # Get the icon path for both .exe and script
    if os.path.exists(icon_path):  # Check if the icon file exists
        try:
            window.iconbitmap(icon_path)  # Set the icon for the window
        except Exception as e:
            print(f"Error setting icon: {e}")
    else:
        print(f"Icon file not found: {icon_path}")  # Print error message if icon file is not found

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        x, y, _, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + cy + 25
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, background="#ffffe0", relief="solid", borderwidth=1, font=("tahoma", "8"))
        label.pack(ipadx=4)

    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

class SerialMonitor:
    def __init__(self, master):
        self.master = master    # Initialize the main window
        self.master.title("FAST-Serial Monitor") # Set window title
        self.master.geometry("1280x720")    # Set initial window size
        self.master.resizable(True, True)   # Allow resizing
        self.master.configure(bg="lightgray")  # Set background color
        self.master.iconbitmap(resource_path("fast.ico"))
        # self.master.attributes("-topmost", True)  # Keep the window on top of other windows
        # self.master.attributes("-alpha", 0.95)  # Set window transparency
        # self.master.attributes("-fullscreen", False)  # Set window to not be fullscreen
        # self.master.attributes("-toolwindow", True)  # Set window to be a tool window (no minimize/maximize buttons)
        # self.master.attributes("-type", "dialog")  # Set window type to dialog (no taskbar button)

        self.captured_messages = []  # List to store message tuples

        self.master.protocol("WM_DELETE_WINDOW", self.on_close)  # Handle window close event

        # Set the window icon (this icon will appear in the title bar)
        set_window_icon(self.master, resource_path("fast.ico"))  # Set the icon for the window
        # Set the icon for the application (this icon will appear in the taskbar)
        icon_path = resource_path("fast.ico")  # Get icon path for both .exe and script

        self.create_widgets()   # Call the method to create widgets

        # Flag to indicate if the serial connection is active
        self.connection_active = False  

        self.manual_header = [ "Serial_Number",
                               "LED_status",
                               "VCC_AVERAGE", 
                               "PLUS_AVERAGE", 
                               "MINUS_AVERAGE",
                               "BATTERY_AVERAGE", 
                               "BUCK_AVERAGE",
                               "VCC",
                               "PLUS_VOLT", 
                               "MINUS_VOLT",
                               "BATTERY_VOLT", 
                               "BUCK_VOLT",
                               "RASPBERRY_PI", 
                               "RASPBERRY_PI_RUN", 
                               "SCREEN_1", 
                               "SCREEN_2", 
                               "SCREEN_3", 
                               "SCREEN_4", 
                               "SCREEN_5",
                               "SCREEN_6",
                                "Status" ]  # Define the manual header for the CSV file
        self.header = self.manual_header  # Set the header to manual header by default

        # Initialize result_dict with "null"
        self.result_dict = {key: "null" for key in self.manual_header}

        # Get the folder where the .exe is running
        base_path = os.path.dirname(os.path.abspath(__file__))

        # Define file paths
        self.csv_path = os.path.join(base_path, 'AM60.csv')
        self.txt_path = os.path.join(base_path, 'captured_messages.txt')

        # Bind close event
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def refresh_ports(self):
        # Clear the current port list and re-populate
        self.port_combobox.set('')  # Clear the current selection
        self.port_combobox_scan.set('')  # Clear the current selection
        self.populate_ports(parent=self.top_frame)  # Re-populate the port combobox with the new list of available ports
        self.populate_ports_scan(parent=self.top_frame)  # Re-populate the scanner port combobox with the new list of available ports

    def on_close(self):
    # Ensure disconnection before closing the app
        if self.connection_active:
            self.disconnect()

        try:
            os.chmod(self.csv_path, 0o666)
            os.remove(self.csv_path)
        except Exception as e:
            print(f"Failed to delete CSV: {e}")

        try:
            os.chmod(self.txt_path, 0o666)
            os.remove(self.txt_path)
        except Exception as e:
            print(f"Failed to delete TXT: {e}")

        self.master.destroy()  # Close the Tkinter window

    def create_widgets(self):  # Method to create the widgets in the main window
        # ========== Set up the grid layout ==========
        # ========== Top Frame (Row 0) - Buttons and Controls ==========
        self.top_frame = ttk.Frame(self.master)
        self.top_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

        # Port and Baud Rate
        self.port_combobox_label = ttk.Label(self.top_frame, text="Select Port:")  
        self.port_combobox_label.grid(row=0, column=0, pady=5)
        ToolTip(self.port_combobox_label, text="COM PORT FOR TEST BANCH")

        self.populate_ports(parent=self.top_frame)  # <-- Pass parent

        self.port_combobox_label_scan = ttk.Label(self.top_frame, text="Scanner Port:")  
        self.port_combobox_label_scan.grid(row=0, column=2, pady=5)
        ToolTip(self.port_combobox_label_scan, text="COM PORT FOR QR SCANNER")

        self.populate_ports_scan(parent=self.top_frame)  # <-- Pass parent

        # Refresh button
        self.refresh_button = ttk.Button(self.top_frame, text="⟳", width=3, command=self.refresh_ports)
        self.refresh_button.grid(row=0, column=4, padx=(10, 10))

        # Baud rate selection
        self.baud_combobox_label = ttk.Label(self.top_frame, width=15, text="Select Baud Rate:")     
        self.baud_combobox_label.grid(row=0, column=5, padx=(5, 0))

        self.baud_combobox = ttk.Combobox(self.top_frame, values=["2400", "4800", "9600", "115200"], width=10, state="readonly")   
        self.baud_combobox.set("115200")    
        self.baud_combobox.grid(row=0, column=6, padx=(5, 5))

        # Buttons
        self.connect_button = ttk.Button(self.top_frame, text="Connect", command=self.connect)
        self.connect_button.grid(row=0, column=7, padx=5)

        self.disconnect_button = ttk.Button(self.top_frame, text="Disconnect", command=self.disconnect, state=tk.DISABLED)
        self.disconnect_button.grid(row=0, column=8, padx=5)

        self.scan_button = ttk.Button(self.top_frame, text="Scan", command=self.scan)
        self.scan_button.grid(row=0, column=9, padx=5)

       # ========== Bottom Frame (Second Row) ========== 
        self.bottom_frame = ttk.Frame(self.master)
        self.bottom_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="nsew")

        self.master.grid_rowconfigure(1, weight=1)
        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_columnconfigure(1, weight=1)

        # Left Column - Log Area
        self.left_frame = ttk.Frame(self.bottom_frame)
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        self.bottom_frame.grid_columnconfigure(0, weight=3)
        self.bottom_frame.grid_rowconfigure(0, weight=1)

        self.log_label = ttk.Label(self.left_frame, text="Scan Result", font=("Helvetica", 10, "bold"))
        self.log_label.grid(row=0, column=0, sticky="w", pady=(0, 2))

        self.log_text = scrolledtext.ScrolledText(self.left_frame, wrap=tk.WORD, height=25, width=90)
        self.log_text.grid(row=1, column=0, sticky="nsew")
        self.left_frame.grid_rowconfigure(1, weight=1)
        self.left_frame.grid_columnconfigure(0, weight=1)

        # Right Column - Previous Message
        self.right_frame = ttk.Frame(self.bottom_frame)
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        self.bottom_frame.grid_columnconfigure(1, weight=2)

        self.latest_label = ttk.Label(self.right_frame, text="Previous Message", font=("Helvetica", 10, "bold"))
        self.latest_label.grid(row=0, column=0, sticky="w", pady=(0, 2))

        self.latest_text = tk.Text(self.right_frame, height=25, width=60, bg="#f0f0f0", wrap="word", state="disabled")
        self.latest_text.grid(row=1, column=0, sticky="nsew")
        self.right_frame.grid_rowconfigure(1, weight=1)
        self.right_frame.grid_columnconfigure(0, weight=1)

    def populate_ports(self, parent):   # Method to populate the combobox with available serial ports
        ports = [port.device for port in serial.tools.list_ports.comports()]    # Get a list of available serial ports
        self.port_combobox = ttk.Combobox(parent, values=ports, state="readonly")  # Create a combobox for selecting ports
        self.port_combobox.grid(row=0, column=1, pady=5, padx=(10, 10))      # Place it in the grid
        ToolTip(self.port_combobox_label, text="COM PORT FOR TEST BANCH")  # Add tooltip to the port label

    def populate_ports_scan(self, parent):   # Method to populate the combobox with available serial ports
        ports = [port.device for port in serial.tools.list_ports.comports()]    # Get a list of available serial ports
        self.port_combobox_scan = ttk.Combobox(parent, values=ports, state="readonly")  # Create a combobox for selecting ports
        self.port_combobox_scan.grid(row=0, column=3, pady=5, padx=(2, 0))      # Place it in the grid
        ToolTip(self.port_combobox_label_scan, text="COM PORT FOR QR SCANNER")  # Add tooltip to the port label

    def connect(self):  # Method to connect to the selected serial port
        port = self.port_combobox.get() # Get the selected port from the combobox
        baud = int(self.baud_combobox.get())    # Get the selected baud rate from the combobox
        try:    
            self.ser = Serial(port, baud, timeout=1)    # Create a Serial object with the selected port and baud rate
            self.log_text.delete(1.0, tk.END)   # Clear the log text area
            self.log_text.insert(tk.END, f"Connected to {port} at {baud} baud\n")   # Insert connection message into the log text area
            self.disconnect_button["state"] = tk.NORMAL # Enable the disconnect button
            self.connect_button["state"] = tk.DISABLED  # Disable the connect button
            self.port_combobox["state"] = tk.DISABLED  # Disable the port combobox
            self.port_combobox_scan["state"] = tk.DISABLED  # Disable the port combobox
            self.refresh_button["state"] = tk.DISABLED      # Disable the refresh button
            self.scan_button["state"] = tk.NORMAL  # Enable the scan button
            self.baud_combobox["state"] = tk.DISABLED  # Enable the refresh button

            self.connection_active = True   # Set the flag to True to indicate connection is active

            self.thread = threading.Thread(target=self.read_from_port)  # Create a thread to read data from the port
            self.thread.start() # Start the thread

        except Exception as e:  # Handle any exceptions that occur during connection
            self.log_text.insert(tk.END, f"Error: {str(e)}\n")  # Insert error message into the log text area

    def disconnect(self):   # Method to disconnect from the serial port
        self.connection_active = False  # Set the flag to False to stop the reading thread
        if hasattr(self, 'ser') and self.ser.is_open:   # Check if the serial object exists and is open
            self.ser.close()    # Close the serial connection

        if hasattr(self, 'thread') and self.thread.is_alive():
            self.thread.join(timeout=1)  # Wait max 2 seconds to prevent freezing

        self.connect_button["state"] = tk.NORMAL    # Enable the connect button
        self.port_combobox["state"] = tk.NORMAL  # Enable the port combobox 
        self.disconnect_button["state"] = tk.DISABLED   # Disable the disconnect button
        self.log_text.insert(tk.END, "Disconnected\n")  # Insert disconnection message into the log text area
        self.log_text.see(tk.END)  # Scroll to the end of the log text area
        self.refresh_button.config(state="enabled")  # Enable the refresh button
        self.baud_combobox.config(state="enabled")  # Enable the baud comebox button
        self.port_combobox_scan["state"] = tk.NORMAL  # Disable the port combobox

    def scan(self):   # Method to disconnect from the serial port
        self.log_text.see(tk.END)  # Scroll to the end of the log text area

        selected_port = self.port_combobox_scan.get()  # Get the selected port for QR scanner
        if selected_port:
            scanned_barcode = barcode_scanner(selected_port)  # Pass the selected port to the scanner function
            if scanned_barcode:  # If barcode was scanned successfully
                self.result_dict = {key: "null" for key in self.manual_header}  # Reset result_dict to "null" for all keys
                self.result_dict["Serial_Number"] = scanned_barcode  # Update with scanned barcode
                dummy_timestamp = "null"  # Use a dummy timestamp
                self.store_data_in_csv(dummy_timestamp, self.result_dict)   # Store the data in CSV

            # If there was a previous valid message, append it to the log
            # This is useful for keeping track of the last scanned barcode
            if hasattr(self, 'last_scanned_barcode'):
                prev_barcode_entry = f"Serial No is : {self.last_scanned_barcode}\n"
                self.last_valid_message = prev_barcode_entry + self.last_valid_message  # Append to previous log

            # move previous valid message to latest_text
            self.latest_text.configure(state='normal')
            self.latest_text.delete("1.0", tk.END)
            self.latest_text.insert(tk.END, self.last_valid_message)  # Insert the last valid message into the latest_text area
            self.latest_text.configure(state='disabled')
            
            # Clear the log text area and insert the scanned barcode
            self.log_text.delete("1.0", tk.END)  # Clear the log text area
            self.log_text.insert(tk.END, f"Serial No is : {scanned_barcode}\n")

            if scanned_barcode:
                self.last_scanned_barcode = scanned_barcode # Store the last scanned barcode
            else:
                self.log_text.insert(tk.END, "No barcode scanned.\n")       
                self.last_scanned_barcode = ""  # Reset last scanned barcode

            

        else:
            self.log_text.insert(tk.END, "No port selected for scanner.\n")

    def read_from_port(self):  # Method to read data from the serial port in a separate thread
        buffer = ""  # Temporary buffer for incoming data
        self.last_valid_message = ""  # Variable to store the last valid message
        start_marker = "START"  # Define start and end markers for messages
        end_marker = "STOP"  # Define start and end markers for messages
        self.log_text.insert(tk.END, "Reading from port...\n")  # Insert reading message into the log text area

        while self.connection_active:
            try:
                line = self.ser.readline().decode("utf-8", errors="ignore")  # Read line
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Get current timestamp
                if line:
                    buffer += line  # Append line to buffer
                    if start_marker in line:
                        self.scan_button["state"] = tk.DISABLED  # Enable the scan button
                        self.log_text.insert(tk.END, f"Time :- {timestamp}\n")  # Show new message in log_text
                    if end_marker in line:
                        self.scan_button["state"] = tk.NORMAL  # Enable the scan button
                        # Show new message in log_text
                    self.log_text.insert(tk.END, line)  # Insert the line into the log text area
                    self.log_text.see(tk.END)  # Scroll to the end of the log text area

                    # Check for complete message between START and END
                    if start_marker in buffer and end_marker in buffer:
                        # Extract the message between START and END markers
                        start_index = buffer.find(start_marker)
                        end_index = buffer.find(end_marker) + len(end_marker)

                        full_message = buffer[start_index:end_index].strip()
                        log_entry = f"[{timestamp}]\n{full_message}\n"

                        if os.path.exists(self.txt_path):
                            os.system(f'attrib -h "{self.txt_path}"')   # Unhide the file after writing
                            os.system(f'attrib -r "{self.txt_path}"')  # Set file permissions to write

                        # Store or log the message
                        with open(self.txt_path, "w") as file:
                            file.write(log_entry)
                        
                        os.system(f'attrib +r "{self.txt_path}"')   # Make the file read-only
                        os.system(f'attrib +h "{self.txt_path}"')   # Hide the file after writing

                        if not self.save_message_exact(full_message, timestamp):
                            buffer = buffer[end_index:]
                            continue  # Skip to the next iteration if the message is not valid

                        self.log_text.see(tk.END)  # Scroll to the end of the log text area

                        # Update last_valid_message to the new message (this is the last good message now)
                        self.last_valid_message = f"[{timestamp}]\n{full_message}"

                        buffer = buffer[end_index:]

            except Exception as e:
                if self.connection_active:
                    self.log_text.insert(tk.END, f"Error reading from port: {str(e)}\n")
                break
 
    def save_message_exact(self, full_message, timestamp):  # Method to save the message to a CSV file

        for key in self.manual_header:
            if key == "Serial_Number" and self.result_dict.get(key, "null") not in ["null", ""]:
            # If the column is "Serial_Number" and it already has a valid value, keep it intact
                continue
            self.result_dict[key] = "null"  # Set all other columns to "null"

        # Split the full message into individual lines
        split_message = [part.strip() for part in full_message.splitlines() if part.strip()]

        # Filter out any comment lines (those starting with /* and ending with */)
        filtered_message = [part for part in split_message if not (part.startswith("/*") and part.endswith("*/"))]

        # If the filtered message is empty (meaning the whole message was a comment), skip it
        if not filtered_message:
            return False  # Signal to skip
        
         # Ensure result_dict exists from scan step
        if not hasattr(self, "result_dict"):
            self.result_dict = {key: "null" for key in self.manual_header}


        # Fixed line-to-header mapping (exact match)
        fixed_lines_map = {
            "LED is not OK": ("LED_status", "0"),
            "LED is OK": ("LED_status", "1"),
            "Vcc Volt is not OK !": ("VCC", "0"),
            "Vcc Volt is OK !": ("VCC", "1"),
            "Plus 5 Volt is not OK": ("PLUS_VOLT", "0"),
            "Plus 5 Volt is OK": ("PLUS_VOLT", "1"),
            "Minus 5 volt is not OK": ("MINUS_VOLT", "0"),
            "Minus 5 volt is OK": ("MINUS_VOLT", "1"),
            "bat is not OK": ("BATTERY_VOLT", "0"),
            "bat is OK": ("BATTERY_VOLT", "1"),
            "buck is not OK": ("BUCK_VOLT", "0"),
            "buck is OK": ("BUCK_VOLT", "1"),
            "Raspberry pi is not working !": ("RASPBERRY_PI", "0"),
            "raspberry pi is OK": ("RASPBERRY_PI", "1"),
            "raspberry pi is running": ("RASPBERRY_PI_RUN", "1"),
        }

        lines = [line.strip() for line in full_message.splitlines() if line.strip()]
        screen_count = 0  # Initialize screen count for each message

        for line in lines:
            if line.startswith("Vcc_avg ="):
                self.result_dict["VCC_AVERAGE"] = line.split("=", 1)[1].strip()
            elif line.startswith("plus_5_avg ="):
                self.result_dict["PLUS_AVERAGE"] = line.split("=", 1)[1].strip()
            elif line.startswith("min_5_avg ="):
                self.result_dict["MINUS_AVERAGE"] = line.split("=", 1)[1].strip()
            elif line.startswith("bat_avg ="):
                self.result_dict["BATTERY_AVERAGE"] = line.split("=", 1)[1].strip()
            elif line.startswith("buck_avg ="):
                self.result_dict["BUCK_AVERAGE"] = line.split("=", 1)[1].strip()
            elif "screen is working ok" in line.lower() or "screen is not ok" in line.lower():          # Check if the line contains "screen is working ok" or "screen is not ok"
                # Assign values to Screen_1, Screen_2, etc., depending on how many screens are being processed
                screen_key = f"SCREEN_{screen_count + 1}"
                self.result_dict[screen_key] = "1" if "working ok" in line.lower() else "0"
                screen_count += 1

            elif line in fixed_lines_map:
                header, value = fixed_lines_map[line]
                self.result_dict[header] = value

        # Error check 1: Serial number is null
        if self.result_dict.get("Serial_Number", "null") == "null":
            print("ERROR: Please scan Serial number")

        # Error check 2: More than 5 nulls from col 3 to col 15
        keys_3_to_15 = self.manual_header[1:16]  # Skip timestamp (assumed separate), take 2nd to 15th headers
        null_count = sum(1 for key in keys_3_to_15 if self.result_dict.get(key, "null") == "null")

        if null_count > 2:
            print("ERROR: More than 2 missing values in diagnostic range (cols 3–15)")
            print("Please Reset Controller and try again")

        # Status check based on LED status and volt statuses (cols 3 and 9–15)
        status_check_keys = [
            "LED status",       # Column 3
            "VCC",              # Column 9
            "PLUS_VOLT",        # Column 10
            "MINUS_VOLT",       # Column 11
            "BATTERY_VOLT",     # Column 12
            "BUCK_VOLT",        # Column 13
            "RASPBERRY_PI",     # Column 14
            "RASPBERRY_PI_RUN",  # Column 15
            "SCREEN_1"          # Column 16
        ]

        status_values = [self.result_dict.get(k, "null") for k in status_check_keys]
        if "0" in status_values:
            self.result_dict["Status"] = "BAD"
        elif all(val == "1" for val in status_values):
            self.result_dict["Status"] = "GOOD"

        # Now, call store_data_in_csv to save this data in CSV
        self.store_data_in_csv(timestamp, self.result_dict)

        # success, message = db_loader.load_csv_to_db()  # Call the function to load data into the database
        # if success:
        #     return True  # Signal success
        # else:
        #     self.log_text.insert(tk.END, f"Database error: {message}\n")
        #     return False  # Signal failure
        return True  # Signal success
        
    def store_data_in_csv(self, timestamp, result_dict): 
        # Use self.result_dict if none is passed
        if result_dict is None:
            result_dict = self.result_dict

        if os.path.exists(self.csv_path):
            os.system(f'attrib -h "{self.csv_path}"')  # Unhide the file before writing'
            os.system(f'attrib -r "{self.csv_path}"')  # Set file permissions to write

        # Always open in write mode to overwrite previous content
        with open(self.csv_path, mode="w", newline="") as file:
            writer = csv.writer(file)
            # If the file doesn't exist, write the header
            writer.writerow(["Timestamp"] + self.manual_header)
            # Then write the new row
            writer.writerow([timestamp] + [result_dict.get(key, "null") for key in self.manual_header])
        
        os.system(f'attrib +r "{self.csv_path}"')  # Make the file read-only
        os.system(f'attrib +h "{self.csv_path}"')   # Hide the file after writing

if __name__ == "__main__":  # Main function to run the application
    root = tk.Tk()  # Create the main window
    # Set the window icon (make sure this is before mainloop)
    set_window_icon(root, "fast.ico")
    app = SerialMonitor(root)   # Create an instance of the SerialMonitor class
    root.mainloop() # Start the Tkinter main loop to run the application
