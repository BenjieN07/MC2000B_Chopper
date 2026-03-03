from thorlabs_mc2000b import MC2000B  # package name
# If your import is different in your environment, try:
# from thorlabs_mc2000b.mc2000b import MC2000B

PORT = "COM4"

def main():
    print(f"Connecting to MC2000B on {PORT} ...")

    ch = MC2000B(serial_port=PORT)  # this will run id? internally and verify it's MC2000B
    print("✅ Connected!")
    print("ID string:", ch.id)

    # Read/print some values to confirm communications are working
    print("\n--- Live readback ---")
    print("Blade:", ch.get_blade_string(), f"(blade index={ch.blade})")
    print("Input reference:", ch.get_inref_string(), f"(ref index={ch.ref})")
    print("Output reference:", ch.get_outref_string(), f"(output index={ch.output})")
    print("Internal frequency (freq):", ch.freq, "Hz")
    print("External input frequency (input):", ch.input, "Hz")
    print("Reference output frequency (refoutfreq):", ch.refoutfreq, "Hz")
    print("Enable (enable):", ch.enable, "(0=still, 1=running)")

    # Optional: toggle enable briefly to prove we can control it (comment out if you don't want motion)
    print("\n--- Control test: toggle enable ---")
    prev = ch.enable
    ch.enable = 1
    print("Set enable=1 ->", ch.enable)
    ch.enable = 0
    print("Set enable=0 ->", ch.enable)
    ch.enable = prev
    print("Restored enable ->", ch.enable)

    ch.close()
    print("\nPort closed. Done.")

if __name__ == "__main__":
    main()