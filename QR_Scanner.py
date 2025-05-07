# from de2120_barcode_scanner import DE2120BarcodeScanner as DE2120
from scanner_lib import DE2120BarcodeScanner as DE2120
import time

# Create a DE2120 object
scanner = DE2120()
buffer = b""

# Wait for and read barcode data
print("Waiting for barcode...")

while True:
    byte = scanner.read()
    if byte:
        if byte in [b'\n', b'\r']:  # End of barcode
            if buffer:
                decoded = buffer.decode('utf-8', errors='ignore').strip()
                print(f"Scanned barcode: {decoded}")
                buffer = b""  # Clear buffer for next scan
        else:
            buffer += byte
    else:
        pass  # No data this moment; wait quietly
