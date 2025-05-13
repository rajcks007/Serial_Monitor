# from de2120_barcode_scanner import DE2120BarcodeScanner as DE2120
from scanner_lib import DE2120BarcodeScanner as DE2120
import time
import serial
from serial.tools import list_ports

def barcode_scanner(port_name):
    # Create a DE2120 object
    scanner = DE2120(port_name)
    buffer = b""  # Initialize an empty byte buffer for barcode data

    scanner.start_scan()  # Start scanning for barcodes

    try:
        while True:
            
            time.sleep(0.1)  # Small delay to allow scanner to scan

            byte = scanner.read()
            if byte:
                time.sleep(0.1)  # Small delay to allow scanner to process data
                if byte in [b'\r', b'\n']:  # End of barcode
                    if buffer:
                        decoded = buffer.decode('utf-8', errors='ignore').strip()
                        print(f"Scanned barcode: {decoded}")
                        buffer = b""  # Clear buffer for next scan
                        scanner.stop_scan()  # Stop scanning
                        return decoded  # Return the scanned barcode
                else:
                    buffer += byte
            else:
                break  # No data this moment; wait quietly
    except serial.SerialException as e:
        print(f"Serial error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")

    return buffer.decode('utf-8', errors='ignore').strip()  # Return the scanned barcode as a string
