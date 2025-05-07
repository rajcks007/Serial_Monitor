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

# Function to get the path to the icon file, works for both script and EXE
def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS  # Path where PyInstaller extracts files
    else:
        base_path = os.path.dirname(__file__)  # Normal Python script path
    return os.path.join(base_path, relative_path)

# Function to set the icon for the Tkinter window
def set_window_icon(window, icon_path):
    if os.path.exists(icon_path):  # Check if the icon file exists
        window.iconbitmap(icon_path)  # Set the icon for the window
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

        self.manual_header = [ "Serial Number",
                               "LED AVERAGE",
                               "VCC AVERAGE", 
                               "PLUS AVERAGE", 
                               "MINUS AVERAGE",
                               "BATTERY AVERAGE",
                               "LED", 
                               "VCC",
                               "PLUS VOLT", 
                               "MINUS VOLT",
                               "BATTERY VOLT", 
                               "RASPBERRY PI", 
                               "RASPBERRY PI RUN", 
                               "SCREEN_1", 
                               "SCREEN_2", 
                               "SCREEN_3", 
                               "SCREEN_4", 
                               "SCREEN_5",
                               "SCREEN_6" ]  # Define the manual header for the CSV file
        self.header = self.manual_header  # Set the header to manual header by default
    
    def refresh_ports(self):
        # Clear the current port list and re-populate
        self.port_combobox.set('')  # Clear the current selection
        self.populate_ports()  # Re-populate the port combobox with the new list of available ports

    def on_close(self):
    # Ensure disconnection before closing the app
        if self.connection_active:
            self.disconnect()
        self.master.destroy()  # Close the Tkinter window

    def create_widgets(self):  # Method to create the widgets in the main window

        # self.export_txt_button = ttk.Button(self.master, text="Export as TXT", command=self.export_txt, state=tk.DISABLED)  # Create a button to export log as TXT
        # self.export_txt_button.grid(row=0, column=6, padx=10, pady=10)  

        # self.export_csv_button = ttk.Button(self.master, text="Export as CSV", command=self.export_csv, state=tk.DISABLED)  # Create a button to export log as CSV
        # self.export_csv_button.grid(row=0, column=7, padx=10, pady=10)  

        # self.export_xml_button = ttk.Button(self.master, text="Export as XML", command=self.export_xml, state=tk.DISABLED)  # Create a button to export log as XML
        # self.export_xml_button.grid(row=0, column=8, padx=10, pady=10)  

        # ========== Main Frames ==========
        self.left_frame = ttk.Frame(self.master)
        self.left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.right_frame = ttk.Frame(self.master)
        self.right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        self.master.grid_columnconfigure(0, weight=3)
        self.master.grid_columnconfigure(1, weight=2)
        self.master.grid_rowconfigure(0, weight=1)

        # ========== Left Frame (Controls + Log) ==========

        # Port and Baud Rate
        self.port_combobox_label = ttk.Label(self.left_frame, text="Select Port:")  
        self.port_combobox_label.grid(row=0, column=0, sticky="w", pady=5)

        self.populate_ports()

        # Refresh Button
        self.refresh_button = ttk.Button(self.left_frame, text="⟳", command=self.refresh_ports)
        self.refresh_button.grid(row=0, column=2, sticky="w", padx=(0, 2))  # Place it in the grid
        ToolTip(self.refresh_button, text="Refresh available ports")  # Add tooltip to the refresh button

        self.baud_combobox_label = ttk.Label(self.left_frame, text="Select Baud Rate:")     
        self.baud_combobox_label.grid(row=0, column=3, sticky="w", pady=5, padx=(2, 0))

        self.baud_combobox = ttk.Combobox(self.left_frame, values=["2400", "4800", "9600", "14400", "115200", "230400", "256000", "460800", "576000", "921600"], state="readonly")   
        self.baud_combobox.set("115200")    
        self.baud_combobox.grid(row=0, column=4, padx=(0, 1))  # Place it in the grid

        # Buttons
        self.connect_button = ttk.Button(self.left_frame, text="Connect", command=self.connect)
        self.connect_button.grid(row=0, column=5, padx=5)

        self.disconnect_button = ttk.Button(self.left_frame, text="Disconnect", command=self.disconnect, state=tk.DISABLED)
        self.disconnect_button.grid(row=0, column=6, padx=5)

        # Log Label
        self.log_label = ttk.Label(self.left_frame, text="Message Log", font=("Helvetica", 10, "bold"))
        self.log_label.grid(row=1, column=0, columnspan=7, sticky="w", pady=(10, 2))

        # Log Text Area
        self.log_text = scrolledtext.ScrolledText(self.left_frame, wrap=tk.WORD, height=25, width=90)
        self.log_text.grid(row=2, column=0, columnspan=7, sticky="nsew", pady=(0, 5))
        self.left_frame.grid_rowconfigure(2, weight=1)

        # ========== Right Frame (Previous Message) ==========

        self.latest_label = ttk.Label(self.right_frame, text="Previous Message", font=("Helvetica", 10, "bold"))
        self.latest_label.grid(row=0, column=0, sticky="w")

        self.latest_text = tk.Text(self.right_frame, height=32, width=60, bg="#f0f0f0", wrap="word", state="disabled")
        self.latest_text.grid(row=1, column=0, sticky="nsew", pady=(5, 0))
        self.right_frame.grid_rowconfigure(1, weight=1)
        self.right_frame.grid_columnconfigure(0, weight=1)


        # Allow columns to expand proportionally
        self.master.grid_columnconfigure(0, weight=3)  # Left wider
        self.master.grid_columnconfigure(1, weight=2)  # Right narrower 

        # Allow row 0 (main content) to expand vertically
        self.master.grid_rowconfigure(0, weight=1)

         # Allow columes to expand vertically
        self.left_frame.grid_columnconfigure(1, weight=1)  # Log area expands vertically
        self.left_frame.grid_columnconfigure(2, weight=1)  # Log area expands vertically 
        self.left_frame.grid_columnconfigure(3, weight=1)  # Log area expands vertically
        self.left_frame.grid_columnconfigure(4, weight=1)  # Log area expands vertically

    def populate_ports(self):   # Method to populate the combobox with available serial ports
        ports = [port.device for port in serial.tools.list_ports.comports()]    # Get a list of available serial ports
        self.port_combobox = ttk.Combobox(self.left_frame, values=ports, state="readonly")  # Create a combobox for selecting ports
        self.port_combobox.grid(row=0, column=1, pady=5, padx=(20, 0))      # Place it in the grid

    def connect(self):  # Method to connect to the selected serial port
        port = self.port_combobox.get() # Get the selected port from the combobox
        baud = int(self.baud_combobox.get())    # Get the selected baud rate from the combobox
        try:    
            self.ser = Serial(port, baud, timeout=1)    # Create a Serial object with the selected port and baud rate
            self.log_text.delete(1.0, tk.END)   # Clear the log text area
            self.log_text.insert(tk.END, f"Connected to {port} at {baud} baud\n")   # Insert connection message into the log text area
            self.disconnect_button["state"] = tk.NORMAL # Enable the disconnect button
            self.connect_button["state"] = tk.DISABLED  # Disable the connect button
            # self.export_txt_button["state"] = tk.NORMAL # Enable the export buttons
            # self.export_csv_button["state"] = tk.NORMAL # Enable the export buttons
            # self.export_xml_button["state"] = tk.NORMAL # Enable the export buttons
            self.port_combobox["state"] = tk.DISABLED  # Disable the port combobox
            self.refresh_button.config(state="disabled")
            self.baud_combobox.config(state="disabled")  # Enable the refresh button


            self.connection_active = True   # Set the flag to True to indicate connection is active

            self.thread = threading.Thread(target=self.read_from_port)  # Create a thread to read data from the port
            self.thread.start() # Start the thread
        except Exception as e:  # Handle any exceptions that occur during connection
            self.log_text.insert(tk.END, f"Error: {str(e)}\n")  # Insert error message into the log text area

    def disconnect(self):   # Method to disconnect from the serial port
        self.connection_active = False  # Set the flag to False to stop the reading thread
        if hasattr(self, 'ser') and self.ser.is_open:   # Check if the serial object exists and is open
            self.ser.close()    # Close the serial connection
        self.connect_button["state"] = tk.NORMAL    # Enable the connect button
        self.port_combobox["state"] = tk.NORMAL  # Enable the port combobox 
        self.disconnect_button["state"] = tk.DISABLED   # Disable the disconnect button
        self.log_text.insert(tk.END, "Disconnected\n")  # Insert disconnection message into the log text area
        self.log_text.see(tk.END)  # Scroll to the end of the log text area
        self.thread.join()  # Wait for the reading thread to finish
        # self.export_txt_button["state"] = tk.DISABLED   # Disable the export buttons
        # self.export_csv_button["state"] = tk.DISABLED   # Disable the export buttons
        # self.export_xml_button["state"] = tk.DISABLED   # Disable the export buttons
        self.log_text.insert(tk.END, "Disconnected\n")  # Insert disconnection message into the log text area
        self.refresh_button.config(state="enabled")  # Enable the refresh button
        self.baud_combobox.config(state="enabled")  # Enable the baud comebox button

    def read_from_port(self):
        buffer = ""  # Temporary buffer for incoming data
        self.last_valid_message = ""  # Variable to store the last valid message
        start_marker = "START"  # Define start and end markers for messages
        end_marker = "STOP"  # Define start and end markers for messages
        self.log_text.insert(tk.END, "Reading from port...\n")  # Insert reading message into the log text area

        while self.connection_active:
            try:
                line = self.ser.readline().decode("utf-8", errors="ignore")  # Read line
                if line:
                    buffer += line  # Append line to buffer

                    # Check for complete message between START and END
                    if start_marker in buffer and end_marker in buffer:
                        # Extract the message between START and END markers
                        start_index = buffer.find(start_marker)
                        end_index = buffer.find(end_marker) + len(end_marker)

                        full_message = buffer[start_index:end_index].strip()
                        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        log_entry = f"[{timestamp}]\n{full_message}\n"

                        # Store or log the message
                        with open("captured_messages.txt", "w") as file:
                            file.write(log_entry)

                        if not self.save_message_exact(full_message, timestamp):
                            buffer = buffer[end_index:]
                            continue  # Skip to the next iteration if the message is not valid

                        # self.log_text.insert(tk.END, log_entry)  # Insert the log entry into the text area
                        self.log_text.see(tk.END)  # Scroll to the end of the log text area

                        # ✅ Only here: move previous valid message to latest_text
                        self.latest_text.configure(state='normal')
                        self.latest_text.delete("1.0", tk.END)
                        self.latest_text.insert(tk.END, self.last_valid_message)  # Insert the last valid message into the latest_text area
                        self.latest_text.configure(state='disabled')

                        # Show new message in log_text
                        self.log_text.delete("1.0", tk.END)
                        self.log_text.insert(tk.END, f"New Message Received:\n{log_entry}")
                        self.log_text.see(tk.END)

                        # Update last_valid_message to the new message (this is the last good message now)
                        self.last_valid_message = f"[{timestamp}]\n{full_message}"

                        # Clear buffer after processing
                        buffer = buffer[end_index:]

            except Exception as e:
                if self.connection_active:
                    self.log_text.insert(tk.END, f"Error reading from port: {str(e)}\n")
                break

    # Save the message to a CSV file with a timestamp
    # This method filters out comment lines and checks for column count before saving
    def save_message_exact(self, full_message, timestamp):
        # Split the full message into individual lines
        split_message = [part.strip() for part in full_message.splitlines() if part.strip()]

        # Filter out any comment lines (those starting with /* and ending with */)
        filtered_message = [part for part in split_message if not (part.startswith("/*") and part.endswith("*/"))]

        # If the filtered message is empty (meaning the whole message was a comment), skip it
        if not filtered_message:
            return False  # Signal to skip
        
        result_dict = {key: "null" for key in self.manual_header}  # Initialize a dictionary with empty values for each header

        # Fixed line-to-header mapping (exact match)
        fixed_lines_map = {
            "LED is not OK": ("LED", "1"),
            "LED is OK": ("LED", "0"),
            "Vcc Volt is not OK !": ("VCC", "1"),
            "Vcc Volt is OK !": ("VCC", "0"),
            "Plus 5 Volt is not OK": ("PLUS VOLT", "1"),
            "Plus 5 Volt is OK": ("PLUS VOLT", "0"),
            "Minus 5 volt is not OK": ("MINUS VOLT", "1"),
            "Minus 5 volt is OK": ("MINUS VOLT", "0"),
            "Battery Volt is not OK !": ("BATTERY VOLT", "1"),
            "Battery Volt is OK !": ("BATTERY VOLT", "0"),
            "Raspberry pi is not working !": ("RASPBERRY PI", "0"),
            "raspberry pi is OK": ("RASPBERRY PI", "1"),
            "raspberry pi is running": ("RASPBERRY PI RUN", "1"),
        }

        lines = [line.strip() for line in full_message.splitlines() if line.strip()]
        screen_count = 0  # Initialize screen count for each message

        for line in lines:
            if line.startswith("led_avg ="):
                result_dict["LED AVERAGE"] = line.split("=", 1)[1].strip()
            elif line.startswith("Vcc_avg ="):
                result_dict["VCC AVERAGE"] = line.split("=", 1)[1].strip()
            elif line.startswith("plus_5_avg ="):
                result_dict["PLUS AVERAGE"] = line.split("=", 1)[1].strip()
            elif line.startswith("min_5_avg ="):
                result_dict["MINUS AVERAGE"] = line.split("=", 1)[1].strip()
            elif line.startswith("Battery_avg ="):
                result_dict["BATTERY AVERAGE"] = line.split("=", 1)[1].strip()
            elif "screen is working ok" in line.lower() or "screen is not ok" in line.lower():          # Check if the line contains "screen is working ok" or "screen is not ok"
                # Assign values to Screen_1, Screen_2, etc., depending on how many screens are being processed
                if screen_count < len([key for key in result_dict if "SCREEN" in key]):
                    # Find the next available "SCREEN_X" slot and assign the value ("1" for "working ok", "0" for "not ok")
                    screen_key = f"SCREEN_{screen_count + 1}"  # This is for Screen_1, Screen_2, etc.
                    result_dict[screen_key] = "1" if "working ok" in line.lower() else "0"
                    screen_count += 1  # Move to the next "SCREEN_X" column

            elif line in fixed_lines_map:
                header, value = fixed_lines_map[line]
                result_dict[header] = value

        # Save to CSV (overwrite to keep only latest message)
        csv_filename = "AM60.csv"
        with open(csv_filename, mode="w", newline="") as csvfile:  # Open in write mode
            writer = csv.writer(csvfile)
            writer.writerow(["Timestamp"] + self.manual_header)
            writer.writerow([timestamp] +  [result_dict[key] for key in self.manual_header])

        return True  # Signal success

        

    """
    # Uncomment the following method if you want to implement TXT export functionality
    def export_txt(self):   # Method to export the log data as a TXT file
        data = self.log_text.get(1.0, tk.END)   # Get all the text from the log text area
        filename = f"serial_log_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.txt" # Create a filename with a timestamp
        with open(filename, "w") as file:   # Open the file in write mode
            file.write(data)    
        self.log_text.insert(tk.END, f"Log exported as TXT: {filename}\n")  # Insert export message into the log text area

    # Uncomment the following method if you want to implement CSV export functionality
    def export_csv(self):   # Method to export the log data as a CSV file
        if not self.captured_messages:
            self.log_text.insert(tk.END, "No data to export!\n")
            return  # Exit early if no data is available
    
        csv_filename = "AM60.csv"  # Create a filename for the CSV file
        file_exists = os.path.isfile(csv_filename)
        # If the file already exists, append to it; otherwise, create a new one
        if not file_exists:
            with open(csv_filename, mode="w", newline="") as csvfile:
                csv_writer = csv.writer(csvfile)
                # Write the header row
                csv_writer.writerow(["Received Time", "Export Time", "Message"])


        with open(csv_filename, mode="a", newline="") as csvfile:
            csv_writer = csv.writer(csvfile)

            # Write the data rows
            for received_time, message in self.captured_messages:
                csv_writer.writerow([received_time, message])

        self.log_text.insert(tk.END, f"Log exported as CSV: {csv_filename}\n")
        self.captured_messages.clear()  # Clear the captured messages list after processing
        
    # Uncomment the following method if you want to implement XML export functionality
    def export_xml(self):   # Method to export the log data as an XML file
        data = self.log_text.get(1.0, tk.END)   # Get all the text from the log text area
        filename = f"serial_log_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xml" # Create a filename with a timestamp
        # Create an XML structure
        root = ET.Element("LogData")    # Create the root element
        lines = data.splitlines()   # Split the data into lines
        for line in lines:  # Iterate through each line
            entry = ET.SubElement(root, "Entry")    # Create a new entry element for each line
            ET.SubElement(entry, "Data").text = line    # Create a sub-element for the data
            ET.SubElement(entry, "Timestamp").text = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        # Write the XML structure to a file
        tree = ET.ElementTree(root)   # Create an ElementTree object
        tree.write(filename)    # Write the XML data to the file
        self.log_text.insert(tk.END, f"Log exported as XML: {filename}\n")  # Insert export message into the log text area
    """

if __name__ == "__main__":  # Main function to run the application
    root = tk.Tk()  # Create the main window
    app = SerialMonitor(root)   # Create an instance of the SerialMonitor class
    root.mainloop() # Start the Tkinter main loop to run the application
