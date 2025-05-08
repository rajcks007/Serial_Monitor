# from de2120_barcode_scanner import DE2120BarcodeScanner as DE2120
from scanner_lib import DE2120BarcodeScanner as DE2120
import time
import serial
import serial.tools.list_ports

# Create a DE2120 object
scanner = DE2120(port_name="COM4")

# Wait for and read barcode data
print("Waiting for barcode...")

def barcode_scanner():
    buffer = b""  # Initialize an empty byte buffer for barcode data

    try:
        while True:
            scanner.start_scan()  # Start scanning for barcodes
            time.sleep(0.1)  # Small delay to allow scanner to scan
            byte = scanner.read()
            if byte:
                if byte in [b'\n', b'\r']:  # End of barcode
                    if buffer:
                        decoded = buffer.decode('utf-8', errors='ignore').strip()
                        print(f"Scanned barcode: {decoded}")
                        buffer = b""  # Clear buffer for next scan
                        break  # Exit the loop after processing the barcode
                else:
                    buffer += byte
            else:
                pass  # No data this moment; wait quietly
    except serial.SerialException as e:
        print(f"Serial error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")
    finally:
        scanner.stop_scan()  # Stop scanning
        if scanner.is_connected():
            scanner.factory_default()  # Reset scanner to factory settings
            print("Scanner reset to factory settings.")

    return decoded  # Return the last scanned barcode