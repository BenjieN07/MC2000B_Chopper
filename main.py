import serial.tools.list_ports

def scan_serial_ports():
    print("Scanning for connected USB/COM devices...\n")
    
    ports = serial.tools.list_ports.comports()
    
    if not ports:
        print("No serial devices found.")
        return
    
    for port in ports:
        print(f"Device: {port.device}")
        print(f"  Description: {port.description}")
        print(f"  Manufacturer: {port.manufacturer}")
        print(f"  HWID: {port.hwid}")
        print("-" * 40)

if __name__ == "__main__":
    scan_serial_ports()