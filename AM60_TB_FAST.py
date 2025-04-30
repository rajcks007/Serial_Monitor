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

        self.master.grid_rowconfigure(0, weight=0)  # Do not allow row 0 to expand
        self.master.grid_rowconfigure(1, weight=1)  # Allow row 1 to expand
        self.master.grid_columnconfigure(0, weight=1)   # Allow column 0 to expand
        self.master.grid_columnconfigure(1, weight=1)   # Allow column 1 to expand
        self.master.grid_columnconfigure(2, weight=1)   # Allow column 2 to expand
        self.master.grid_columnconfigure(3, weight=1)   # Allow column 3 to expand
        self.master.grid_columnconfigure(4, weight=1)   # Allow column 4 to expand
        self.master.grid_columnconfigure(5, weight=1)   # Allow column 5 to expand
        self.master.grid_columnconfigure(6, weight=1)   # Allow column 6 to expand
        self.master.grid_columnconfigure(7, weight=1)   # Allow column 7 to expand  
        self.master.grid_columnconfigure(8, weight=1)   # Allow column 8 to expand
        self.master.grid_rowconfigure(0, weight=0)  # Do not allow row 0 to expand
        self.master.grid_rowconfigure(1, weight=1)  # Allow row 1 to expand

        self.master.protocol("WM_DELETE_WINDOW", self.on_close)  # Handle window close event

        # Set the window icon (this icon will appear in the title bar)
        set_window_icon(self.master, resource_path("fast.ico"))  # Set the icon for the window
        # Set the icon for the application (this icon will appear in the taskbar)
        icon_path = resource_path("fast.ico")  # Get icon path for both .exe and script

        self.create_widgets()   # Call the method to create widgets

        # Flag to indicate if the serial connection is active
        self.connection_active = False  

    def on_close(self):
    # Ensure disconnection before closing the app
        if self.connection_active:
            self.disconnect()
        self.master.destroy()  # Close the Tkinter window


    def create_widgets(self):
        self.port_combobox_label = ttk.Label(self.master, text="Select Port:")  
        self.port_combobox_label.grid(row=0, column=0, padx=10, pady=10)        

        self.populate_ports()   # Populate the combobox with available serial ports

        self.baud_combobox_label = ttk.Label(self.master, text="Select Baud Rate:")     
        self.baud_combobox_label.grid(row=0, column=2, padx=10, pady=10)        

        self.baud_combobox = ttk.Combobox(self.master, values=["2400","4800","9600","14400", "115200"], state="readonly")   # Create a combobox for baud rates
        self.baud_combobox.set("115200")    # Set default baud rate
        self.baud_combobox.grid(row=0, column=3, padx=10, pady=10)      

        self.connect_button = ttk.Button(self.master, text="Connect", command=self.connect) # Create a button to connect to the selected port
        self.connect_button.grid(row=0, column=4, padx=10, pady=10)     

        self.disconnect_button = ttk.Button(self.master, text="Disconnect", command=self.disconnect, state=tk.DISABLED) # Create a button to disconnect from the port
        self.disconnect_button.grid(row=0, column=5, padx=10, pady=10)

        # self.export_txt_button = ttk.Button(self.master, text="Export as TXT", command=self.export_txt, state=tk.DISABLED)  # Create a button to export log as TXT
        # self.export_txt_button.grid(row=0, column=6, padx=10, pady=10)  

        self.export_csv_button = ttk.Button(self.master, text="Export as CSV", command=self.export_csv, state=tk.DISABLED)  # Create a button to export log as CSV
        self.export_csv_button.grid(row=0, column=7, padx=10, pady=10)  

        # self.export_xml_button = ttk.Button(self.master, text="Export as XML", command=self.export_xml, state=tk.DISABLED)  # Create a button to export log as XML
        # self.export_xml_button.grid(row=0, column=8, padx=10, pady=10)  

        self.log_text = scrolledtext.ScrolledText(self.master, wrap=tk.WORD)    # Create a scrolled text area for displaying log data
        self.log_text.grid(row=1, column=0, columnspan=9, padx=10, pady=10, sticky="nsew")       # Expand to fill the window

    def populate_ports(self):   # Method to populate the combobox with available serial ports
        ports = [port.device for port in serial.tools.list_ports.comports()]    # Get a list of available serial ports
        self.port_combobox = ttk.Combobox(self.master, values=ports, state="readonly")  # Create a combobox for selecting ports
        self.port_combobox.grid(row=0, column=1, padx=10, pady=10)      # Place it in the grid

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
            self.export_csv_button["state"] = tk.NORMAL # Enable the export buttons
            # self.export_xml_button["state"] = tk.NORMAL # Enable the export buttons
            self.port_combobox["state"] = tk.DISABLED  # Disable the port combobox

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
        self.export_csv_button["state"] = tk.DISABLED   # Disable the export buttons
        # self.export_xml_button["state"] = tk.DISABLED   # Disable the export buttons
        self.log_text.insert(tk.END, "Disconnected\n")  # Insert disconnection message into the log text area

    def read_from_port(self):
        buffer = ""  # Temporary buffer for incoming data
        start_marker = "START"  # Define start and end markers for messages
        end_marker = "STOP"  # Define start and end markers for messages
        self.log_text.insert(tk.END, "Reading from port...\n")  # Insert reading message into the log text area

        while self.connection_active:
            try:
                line = self.ser.readline().decode("utf-8", errors="ignore")  # Read line
                if line:
                    buffer += line  # Append line to buffer
                    # Insert the line into the log text area
                    self.log_text.insert(tk.END, line)  # Insert the line into the log text area
                    self.log_text.see(tk.END)   # Scroll to the end of the log text area

                    # Check for complete message between START and END
                    if start_marker in buffer and end_marker in buffer:
                        # Extract the message between START and END markers
                        start_index = buffer.find(start_marker)
                        end_index = buffer.find(end_marker) + len(end_marker)

                        full_message = buffer[start_index:end_index].strip()
                        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        log_entry = f"[{timestamp}] {full_message}\n"

                        # Store or log the message
                        with open("captured_messages.txt", "w") as file:
                            file.write(log_entry)

                        self.log_text.insert(tk.END, log_entry)  # Insert the log entry into the text area
                        self.log_text.see(tk.END)  # Scroll to the end of the log text area

                        self.captured_messages.append((timestamp, full_message))  # Store the message with timestamp

                        # Clear screen and show the extracted message
                        self.log_text.delete(1.0, tk.END)
                        self.log_text.insert(tk.END, f"New Message Received:\n{log_entry}")
                        self.log_text.see(tk.END)

                        # Clear buffer after processing
                        buffer = buffer[end_index:]

            except Exception as e:
                if self.connection_active:
                    self.log_text.insert(tk.END, f"Error reading from port: {str(e)}\n")
                break
        

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
