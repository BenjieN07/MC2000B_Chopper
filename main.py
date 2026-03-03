import serial

PORT = "COM4"      # your chopper port
BAUD = 115200
TIMEOUT = 5.0

def query(ser: serial.Serial, cmd: str, read_n: int = 64) -> bytes:
    """
    Send command terminated with CR and read back a fixed number of bytes.
    Returns raw bytes so you can debug what the device actually sent.
    """
    if not cmd.endswith("\r"):
        cmd += "\r"
    ser.reset_input_buffer()
    ser.write(cmd.encode("ascii"))
    resp = ser.read(read_n)
    return resp

def parse_numeric(resp: bytes) -> float:
    """
    Extract the first number found in the response.
    This is robust against extra chars like prompts, CR/LF, etc.
    """
    text = resp.decode(errors="ignore")
    # Keep digits, decimal point, sign; split on anything else
    token = ""
    for ch in text:
        if ch.isdigit() or ch in ".-+":
            token += ch
        elif token:
            break
    if not token:
        raise ValueError(f"Could not parse number from response: {text!r}")
    return float(token)

def main():
    print(f"Connecting to {PORT} @ {BAUD}...")
    with serial.Serial(
        port=PORT,
        baudrate=BAUD,
        timeout=TIMEOUT,
        bytesize=serial.EIGHTBITS,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        xonxoff=False,
        rtscts=False,
        dsrdtr=False,
    ) as ser:

        # Same “wake up” trick as the GitHub code
        ser.write(b"\r")
        _ = ser.read(100)

        # --- Queries (same commands as chopper_controller.py) ---
        # Internal frequency
        raw = query(ser, "freq?", read_n=64)
        intfreq = parse_numeric(raw)
        print(f"freq?   raw={raw!r}")
        print(f"Internal frequency: {intfreq:.3f} Hz\n")

        # Running/still status (enable? -> 0 or 1)
        raw = query(ser, "enable?", read_n=64)
        enable = parse_numeric(raw)
        print(f"enable? raw={raw!r}")
        print(f"Enable status: {int(enable)} (0=still, 1=running)\n")

        # Blade index/type (blade?)
        raw = query(ser, "blade?", read_n=64)
        blade = parse_numeric(raw)
        print(f"blade?  raw={raw!r}")
        print(f"Blade index: {int(blade)}\n")

        # Reference mode (ref?) 0=internal, 1=external (per the GitHub code design)
        raw = query(ser, "ref?", read_n=64)
        ref = parse_numeric(raw)
        print(f"ref?    raw={raw!r}")
        print(f"Reference mode: {int(ref)} (0=internal, 1=external)\n")

        # External reference input frequency (input?)
        raw = query(ser, "input?", read_n=64)
        exfreq = parse_numeric(raw)
        print(f"input?  raw={raw!r}")
        print(f"External input frequency: {exfreq:.3f} Hz\n")

    print("Done. (Port closed)")

if __name__ == "__main__":
    main()