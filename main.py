import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

from thorlabs_mc2000b import MC2000B


DEFAULT_PORT = "COM4"


class MC2000BCleanGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MC2000B Control Utility (Clean)")
        self.geometry("980x560")
        self.minsize(980, 560)

        self.ch = None
        self.poll_ms = 400
        self._poll_job = None

        # Toolbar
        self.port_var = tk.StringVar(value=DEFAULT_PORT)
        self.status_var = tk.StringVar(value="Disconnected")
        self.id_var = tk.StringVar(value="—")

        # Left readouts
        self.actual_freq_var = tk.StringVar(value="—")    # panel-like
        self.target_freq_var = tk.StringVar(value="—")    # setpoint
        self.ext_in_var = tk.StringVar(value="—")         # REF IN
        self.ref_out_var = tk.StringVar(value="—")        # REF OUT

        # Status box
        self.enable_text_var = tk.StringVar(value="—")
        self.blade_text_var = tk.StringVar(value="—")
        self.inref_text_var = tk.StringVar(value="—")
        self.outref_text_var = tk.StringVar(value="—")

        # Controls
        self.output_sel = tk.StringVar(value="inner")              # target/inner/outer
        self.ref_sel = tk.StringVar(value="internal-inner")        # internal-outer/internal-inner/external-outer/external-inner

        self.nharm_var = tk.IntVar(value=1)
        self.dharm_var = tk.IntVar(value=1)

        self.phase_var = tk.IntVar(value=180)
        self.target_set_var = tk.IntVar(value=305)                 # user setpoint on right
        self.cycle_var = tk.IntVar(value=1)

        # Console
        self.console = None

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------------- UI ----------------
    def _build_ui(self):
        self.configure(bg="white")
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(".", background="white", foreground="#111", fieldbackground="white")
        style.configure("Panel.TFrame", background="#f4f6f8")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Title.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("Big.TLabel", font=("Segoe UI", 28, "bold"))
        style.configure("Unit.TLabel", font=("Segoe UI", 11, "bold"), foreground="#555")
        style.configure("Muted.TLabel", foreground="#555")
        style.configure("Accent.TButton", background="#1f6feb", foreground="white")
        style.map("Accent.TButton", background=[("active", "#3b82f6")])

        # ---------- Top bar ----------
        top = ttk.Frame(self, style="Panel.TFrame")
        top.pack(fill="x", padx=10, pady=(10, 8))

        ttk.Label(top, text="Device", style="Title.TLabel", background="#f4f6f8").pack(side="left", padx=(10, 14))
        ttk.Label(top, text="Port:", background="#f4f6f8").pack(side="left")
        ttk.Entry(top, textvariable=self.port_var, width=10).pack(side="left", padx=(6, 12))

        ttk.Button(top, text="Connect", style="Accent.TButton", command=self.connect).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Disconnect", command=self.disconnect).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Refresh", command=lambda: self.refresh_once(log_it=True)).pack(side="left", padx=(0, 8))

        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=10, pady=8)
        ttk.Label(top, textvariable=self.status_var, background="#f4f6f8").pack(side="left")
        ttk.Label(top, textvariable=self.id_var, style="Muted.TLabel", background="#f4f6f8").pack(side="right", padx=10)

        # ---------- Main ----------
        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        left = ttk.Frame(main, style="Panel.TFrame")
        left.pack(side="left", fill="y", padx=(0, 10))
        left.configure(width=360)
        left.pack_propagate(False)

        right = ttk.Frame(main, style="Panel.TFrame")
        right.pack(side="right", fill="both", expand=True)

        # Left readout cards
        self._readout_card(left, "Actual Frequency (Hz) — panel-like (0 when stopped)", self.actual_freq_var, "Hz", big=True)
        self._readout_card(left, "Target Frequency (Hz) — setpoint", self.target_freq_var, "Hz", big=False)
        self._readout_card(left, "External Input Frequency (Hz) — REF IN", self.ext_in_var, "Hz", big=False)
        self._readout_card(left, "Ref-Out Frequency (Hz) — REF OUT", self.ref_out_var, "Hz", big=False)

        status_box = ttk.LabelFrame(left, text="Status")
        status_box.pack(fill="x", padx=10, pady=(8, 10))

        self._kv(status_box, "Enable:", self.enable_text_var, 0)
        self._kv(status_box, "Blade:", self.blade_text_var, 1)
        self._kv(status_box, "In Ref:", self.inref_text_var, 2)
        self._kv(status_box, "Out Ref:", self.outref_text_var, 3)

        # Right side layout using PACK so Start/Stop never disappears
        right_inner = ttk.Frame(right, style="Panel.TFrame")
        right_inner.pack(fill="both", expand=True, padx=10, pady=10)

        # Row 1: Output + Reference on left, sliders on right
        top_row = ttk.Frame(right_inner, style="Panel.TFrame")
        top_row.pack(fill="x")

        left_controls = ttk.Frame(top_row, style="Panel.TFrame")
        left_controls.pack(side="left", fill="y", padx=(0, 10))

        sliders = ttk.Frame(top_row, style="Panel.TFrame")
        sliders.pack(side="left", fill="x", expand=True)

        # Output
        out_box = ttk.LabelFrame(left_controls, text="Output (REF OUT source)")
        out_box.pack(fill="x", pady=(0, 10))

        ttk.Radiobutton(out_box, text="Target", variable=self.output_sel, value="target",
                        command=self.apply_output).grid(row=0, column=0, sticky="w", padx=10, pady=6)
        ttk.Radiobutton(out_box, text="Inner", variable=self.output_sel, value="inner",
                        command=self.apply_output).grid(row=1, column=0, sticky="w", padx=10, pady=6)
        ttk.Radiobutton(out_box, text="Outer", variable=self.output_sel, value="outer",
                        command=self.apply_output).grid(row=0, column=1, sticky="w", padx=10, pady=6)

        # Reference
        ref_box = ttk.LabelFrame(left_controls, text="Reference (lock source)")
        ref_box.pack(fill="x")

        ttk.Radiobutton(ref_box, text="INT Outer", variable=self.ref_sel, value="internal-outer",
                        command=self.apply_inref).grid(row=0, column=0, sticky="w", padx=10, pady=6)
        ttk.Radiobutton(ref_box, text="INT Inner", variable=self.ref_sel, value="internal-inner",
                        command=self.apply_inref).grid(row=0, column=1, sticky="w", padx=10, pady=6)
        ttk.Radiobutton(ref_box, text="EXT Outer", variable=self.ref_sel, value="external-outer",
                        command=self.apply_inref).grid(row=1, column=0, sticky="w", padx=10, pady=6)
        ttk.Radiobutton(ref_box, text="EXT Inner", variable=self.ref_sel, value="external-inner",
                        command=self.apply_inref).grid(row=1, column=1, sticky="w", padx=10, pady=6)

        md = ttk.Frame(ref_box)
        md.grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=(8, 10))

        ttk.Label(md, text="Multiplier").grid(row=0, column=0, sticky="w")
        ttk.Entry(md, textvariable=self.nharm_var, width=6).grid(row=1, column=0, sticky="w", padx=(0, 6))
        ttk.Label(md, text="/").grid(row=1, column=1, sticky="w")
        ttk.Label(md, text="Divider").grid(row=0, column=2, sticky="w", padx=(10, 0))
        ttk.Entry(md, textvariable=self.dharm_var, width=6).grid(row=1, column=2, sticky="w", padx=(0, 6))
        ttk.Button(md, text="Apply", command=self.apply_harmonics).grid(row=1, column=3, sticky="w", padx=(8, 0))

        # Sliders (right side)
        self._slider_block(sliders, title="Phase (Degrees)", var=self.phase_var, from_=0, to=360, apply_cmd=self.apply_phase)
        self._slider_block(sliders, title="Target Frequency (Hz)", var=self.target_set_var, from_=1, to=10000, apply_cmd=self.apply_target_freq)
        self._slider_block(sliders, title="Cycle (%)", var=self.cycle_var, from_=1, to=50, apply_cmd=self.apply_cycle)

        # Row 2: Start/Stop (always visible)
        run_row = ttk.Frame(right_inner, style="Panel.TFrame")
        run_row.pack(fill="x", pady=(10, 6))

        ttk.Button(run_row, text="Start", style="Accent.TButton", command=self.start).pack(side="left", padx=(0, 8))
        ttk.Button(run_row, text="Stop", command=self.stop).pack(side="left")

        # Row 3: Console
        console_box = ttk.LabelFrame(right_inner, text="Console")
        console_box.pack(fill="both", expand=True)

        self.console = tk.Text(console_box, height=10, wrap="word",
                               bg="white", fg="#111", insertbackground="#111")
        self.console.pack(fill="both", expand=True, padx=10, pady=10)
        self._log("Ready. Connect to COM4.\n")

    def _readout_card(self, parent, title, var, unit, big=False):
        card = ttk.Frame(parent, style="Card.TFrame")
        card.pack(fill="x", padx=10, pady=(10, 0))

        ttk.Label(card, text=title, style="Muted.TLabel", background="#ffffff").pack(anchor="w", padx=10, pady=(10, 0))

        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill="x", padx=10, pady=(0, 10))

        if big:
            ttk.Label(row, textvariable=var, style="Big.TLabel", background="#ffffff").pack(side="left")
        else:
            ttk.Label(row, textvariable=var, font=("Segoe UI", 18, "bold"), background="#ffffff").pack(side="left")
        ttk.Label(row, text=unit, style="Unit.TLabel", background="#ffffff").pack(side="left", padx=(10, 0))

    def _kv(self, parent, k, v_var, row):
        ttk.Label(parent, text=k).grid(row=row, column=0, sticky="w", padx=10, pady=4)
        ttk.Label(parent, textvariable=v_var, style="Muted.TLabel").grid(row=row, column=1, sticky="w", padx=10, pady=4)

    def _slider_block(self, parent, title, var, from_, to, apply_cmd):
        box = ttk.LabelFrame(parent, text=title)
        box.pack(fill="x", pady=(0, 10))

        top = ttk.Frame(box)
        top.pack(fill="x", padx=10, pady=(8, 4))

        entry = ttk.Entry(top, width=8)
        entry.insert(0, str(var.get()))
        entry.pack(side="right")
        ttk.Button(top, text="Apply", command=lambda e=entry: apply_cmd(e.get())).pack(side="right", padx=(0, 8))

        s = ttk.Scale(box, from_=from_, to=to, orient="horizontal",
                      command=lambda _v, e=entry: self._sync_entry_from_scale(e, _v))
        s.set(var.get())
        s.pack(fill="x", padx=10, pady=(0, 10))

    def _sync_entry_from_scale(self, entry, v):
        try:
            val = int(float(v))
        except Exception:
            return
        entry.delete(0, "end")
        entry.insert(0, str(val))

    # ---------- logging ----------
    def _log(self, msg: str):
        self.console.insert("end", msg)
        self.console.see("end")

    def _ts(self):
        return datetime.now().strftime("%H:%M:%S")

    # ---------- polling ----------
    def _start_poll(self):
        self._stop_poll()
        self._poll_job = self.after(self.poll_ms, self._poll_tick)

    def _stop_poll(self):
        if self._poll_job is not None:
            try:
                self.after_cancel(self._poll_job)
            except Exception:
                pass
        self._poll_job = None

    def _poll_tick(self):
        self.refresh_once(log_it=False)
        self._poll_job = self.after(self.poll_ms, self._poll_tick)

    # ---------- connect/disconnect ----------
    def connect(self):
        if self.ch is not None:
            return
        port = self.port_var.get().strip()
        try:
            self.ch = MC2000B(serial_port=port)
            self.status_var.set(f"Connected ({port})")
            self.id_var.set(self.ch.id)
            self._log(f"[{self._ts()}] Connected: {self.ch.id}\n")
            self._sync_controls_from_device()
            self.refresh_once(log_it=False)
            self._start_poll()
        except Exception as e:
            self.ch = None
            self.status_var.set("Disconnected")
            self.id_var.set("—")
            messagebox.showerror("Connect failed", str(e))
            self._log(f"[{self._ts()}] ERROR connect: {e}\n")

    def disconnect(self):
        self._stop_poll()
        if self.ch is not None:
            try:
                self.ch.close()
            except Exception:
                pass
        self.ch = None
        self.status_var.set("Disconnected")
        self.id_var.set("—")
        self._log(f"[{self._ts()}] Disconnected\n")

        for v in (
            self.actual_freq_var, self.target_freq_var, self.ext_in_var, self.ref_out_var,
            self.enable_text_var, self.blade_text_var, self.inref_text_var, self.outref_text_var
        ):
            v.set("—")

    def on_close(self):
        self.disconnect()
        self.destroy()

    # ---------- sync ----------
    def _sync_controls_from_device(self):
        if self.ch is None:
            return
        try:
            self.output_sel.set(self.ch.get_outref_string())
            self.ref_sel.set(self.ch.get_inref_string())
            self.phase_var.set(int(self.ch.phase))
            self.target_set_var.set(int(self.ch.freq))
            self.cycle_var.set(int(self.ch.oncycle))
            self.nharm_var.set(int(self.ch.nharmonic))
            self.dharm_var.set(int(self.ch.dharmonic))
        except Exception as e:
            self._log(f"[{self._ts()}] WARN sync: {e}\n")

    # ---------- refresh ----------
    def refresh_once(self, log_it: bool):
        if self.ch is None:
            return
        try:
            enable = int(self.ch.enable)
            target = int(self.ch.freq)
            refout = int(self.ch.refoutfreq)
            extin = int(self.ch.input)
            if extin < 0:
                extin = 0

            actual_panel = 0 if enable == 0 else (refout if refout > 0 else target)

            self.actual_freq_var.set(str(actual_panel))
            self.target_freq_var.set(str(target))
            self.ext_in_var.set(str(extin))
            self.ref_out_var.set(str(refout))

            self.enable_text_var.set("RUNNING" if enable == 1 else "STOPPED")
            self.blade_text_var.set(self.ch.get_blade_string())
            self.inref_text_var.set(self.ch.get_inref_string())
            self.outref_text_var.set(self.ch.get_outref_string())

            # keep buttons synced
            self.output_sel.set(self.ch.get_outref_string())
            self.ref_sel.set(self.ch.get_inref_string())

            if log_it:
                self._log(f"[{self._ts()}] Refresh: enable={enable}, target={target}, refout={refout}, extin={extin}\n")
        except Exception as e:
            self._log(f"[{self._ts()}] ERROR refresh: {e}\n")

    # ---------- apply controls ----------
    def apply_output(self):
        if self.ch is None:
            return
        try:
            val = self.output_sel.get().strip()
            self.ch.set_outref_string(val)
            self._log(f"[{self._ts()}] Set output='{val}'\n")
            self.refresh_once(log_it=False)
        except Exception as e:
            messagebox.showerror("Set Output failed", str(e))
            self._log(f"[{self._ts()}] ERROR output: {e}\n")

    def apply_inref(self):
        if self.ch is None:
            return
        try:
            val = self.ref_sel.get().strip()
            self.ch.set_inref_string(val)
            self._log(f"[{self._ts()}] Set reference='{val}'\n")
            self.refresh_once(log_it=False)
        except Exception as e:
            messagebox.showerror("Set Reference failed", str(e))
            self._log(f"[{self._ts()}] ERROR reference: {e}\n")

    def apply_harmonics(self):
        if self.ch is None:
            return
        try:
            self.ch.nharmonic = int(self.nharm_var.get())
            self.ch.dharmonic = int(self.dharm_var.get())
            self._log(f"[{self._ts()}] Set nharmonic={self.ch.nharmonic}, dharmonic={self.ch.dharmonic}\n")
            self.refresh_once(log_it=False)
        except Exception as e:
            messagebox.showerror("Set Mult/Div failed", str(e))
            self._log(f"[{self._ts()}] ERROR harmonics: {e}\n")

    def apply_phase(self, value):
        if self.ch is None:
            return
        try:
            v = int(float(value))
            self.ch.phase = v
            self.phase_var.set(v)
            self._log(f"[{self._ts()}] Set phase={v}\n")
            self.refresh_once(log_it=False)
        except Exception as e:
            messagebox.showerror("Set Phase failed", str(e))
            self._log(f"[{self._ts()}] ERROR phase: {e}\n")

    def apply_target_freq(self, value):
        if self.ch is None:
            return
        try:
            v = int(float(value))
            self.ch.freq = v
            self.target_set_var.set(v)
            self._log(f"[{self._ts()}] Set TARGET freq={v}\n")
            self.refresh_once(log_it=False)  # updates left target immediately
        except Exception as e:
            messagebox.showerror("Set Target Frequency failed", str(e))
            self._log(f"[{self._ts()}] ERROR target freq: {e}\n")

    def apply_cycle(self, value):
        if self.ch is None:
            return
        try:
            v = int(float(value))
            self.ch.oncycle = v
            self.cycle_var.set(v)
            self._log(f"[{self._ts()}] Set cycle={v}\n")
            self.refresh_once(log_it=False)
        except Exception as e:
            messagebox.showerror("Set Cycle failed", str(e))
            self._log(f"[{self._ts()}] ERROR cycle: {e}\n")

    def start(self):
        if self.ch is None:
            return
        try:
            self.ch.enable = 1
            self._log(f"[{self._ts()}] Start (enable=1)\n")
            self.refresh_once(log_it=False)
        except Exception as e:
            messagebox.showerror("Start failed", str(e))
            self._log(f"[{self._ts()}] ERROR start: {e}\n")

    def stop(self):
        if self.ch is None:
            return
        try:
            self.ch.enable = 0
            self._log(f"[{self._ts()}] Stop (enable=0)\n")
            self.refresh_once(log_it=False)
        except Exception as e:
            messagebox.showerror("Stop failed", str(e))
            self._log(f"[{self._ts()}] ERROR stop: {e}\n")


if __name__ == "__main__":
    MC2000BCleanGUI().mainloop()