import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

from thorlabs_mc2000b import MC2000B, Blade

DEFAULT_PORT = "COM4"


def set_light_style(root: tk.Tk):
    style = ttk.Style(root)
    style.theme_use("clam")

    bg = "#ffffff"
    panel = "#f4f6f8"
    panel2 = "#ffffff"
    fg = "#111111"
    muted = "#555555"
    accent = "#1f6feb"   # blue
    ok = "#14804a"       # green
    warn = "#b45309"     # amber
    bad = "#b42318"      # red

    root.configure(background=bg)

    style.configure(".", background=bg, foreground=fg, fieldbackground=panel2)
    style.configure("TFrame", background=bg)
    style.configure("Panel.TFrame", background=panel)
    style.configure("Panel2.TFrame", background=panel2)

    style.configure("TLabel", background=bg, foreground=fg)
    style.configure("Muted.TLabel", foreground=muted)
    style.configure("Title.TLabel", font=("Segoe UI", 12, "bold"))
    style.configure("Big.TLabel", font=("Segoe UI", 26, "bold"))
    style.configure("BigUnit.TLabel", font=("Segoe UI", 11, "bold"), foreground=muted)

    style.configure("TButton", padding=(10, 6))
    style.configure("Accent.TButton", background=accent, foreground="#ffffff")
    style.map("Accent.TButton", background=[("active", "#3b82f6")])

    style.configure("TEntry", padding=6)
    style.configure("TCombobox", padding=6)

    style.configure("TLabelframe", background=bg, foreground=fg)
    style.configure("TLabelframe.Label", background=bg, foreground=fg, font=("Segoe UI", 10, "bold"))

    style.configure("TNotebook", background=bg, borderwidth=0)
    style.configure("TNotebook.Tab", background=panel, foreground=fg, padding=(12, 8))
    style.map("TNotebook.Tab",
              background=[("selected", panel2)],
              foreground=[("selected", fg)])

    style.configure("StatusOK.TLabel", background=bg, foreground=ok, font=("Segoe UI", 10, "bold"))
    style.configure("StatusBAD.TLabel", background=bg, foreground=bad, font=("Segoe UI", 10, "bold"))
    style.configure("StatusWARN.TLabel", background=bg, foreground=warn, font=("Segoe UI", 10, "bold"))

    return {"bg": bg, "panel": panel, "panel2": panel2, "fg": fg, "muted": muted,
            "accent": accent, "ok": ok, "warn": warn, "bad": bad}


class MC2000BApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MC2000B Control Utility")
        self.geometry("880x560")
        self.minsize(880, 560)

        self.colors = set_light_style(self)

        self.ch = None
        self.auto_refresh_ms = 500
        self._refresh_job = None

        # toolbar vars
        self.port_var = tk.StringVar(value=DEFAULT_PORT)
        self.status_text = tk.StringVar(value="Disconnected")

        # readback vars
        self.id_var = tk.StringVar(value="—")
        self.freq_rb = tk.StringVar(value="—")
        self.input_rb = tk.StringVar(value="—")
        self.refout_rb = tk.StringVar(value="—")
        self.blade_rb = tk.StringVar(value="—")
        self.inref_rb = tk.StringVar(value="—")
        self.outref_rb = tk.StringVar(value="—")
        self.enable_rb = tk.StringVar(value="—")

        # control vars
        self.enable_set = tk.IntVar(value=0)
        self.freq_set = tk.StringVar(value="200")
        self.nharm_set = tk.StringVar(value="1")
        self.dharm_set = tk.StringVar(value="1")
        self.phase_set = tk.StringVar(value="0")
        self.oncycle_set = tk.StringVar(value="50")
        self.intensity_set = tk.StringVar(value="100")

        self.blade_choice = tk.StringVar(value="")
        self.inref_choice = tk.StringVar(value="")
        self.outref_choice = tk.StringVar(value="")

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------------- UI ----------------
    def _build_ui(self):
        # Top toolbar
        toolbar = ttk.Frame(self, style="Panel.TFrame")
        toolbar.pack(fill="x", padx=10, pady=(10, 8))

        ttk.Label(toolbar, text="Device", style="Title.TLabel", background=self.colors["panel"]).pack(side="left", padx=(10, 16))

        ttk.Label(toolbar, text="Port:", background=self.colors["panel"]).pack(side="left")
        ttk.Entry(toolbar, textvariable=self.port_var, width=10).pack(side="left", padx=(6, 12))

        ttk.Button(toolbar, text="Connect", style="Accent.TButton", command=self.connect).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Disconnect", command=self.disconnect).pack(side="left", padx=(0, 8))
        ttk.Button(toolbar, text="Refresh", command=self.refresh_once).pack(side="left", padx=(0, 8))

        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=10, pady=8)

        self.status_label = ttk.Label(toolbar, textvariable=self.status_text, style="StatusBAD.TLabel", background=self.colors["panel"])
        self.status_label.pack(side="left")

        ttk.Label(toolbar, textvariable=self.id_var, style="Muted.TLabel", background=self.colors["panel"]).pack(side="right", padx=10)

        # Main split
        main = ttk.Frame(self)
        main.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        left = ttk.Frame(main, style="Panel.TFrame")
        left.pack(side="left", fill="y", padx=(0, 10))
        left.configure(width=320)
        left.pack_propagate(False)

        right = ttk.Frame(main, style="Panel.TFrame")
        right.pack(side="right", fill="both", expand=True)

        # Left big readouts
        self._readout_card(left, "Frequency", self.freq_rb, "Hz")
        self._readout_card(left, "External Input", self.input_rb, "Hz")
        self._readout_card(left, "Ref-Out", self.refout_rb, "Hz")

        meta = ttk.LabelFrame(left, text="Status")
        meta.pack(fill="x", padx=10, pady=(8, 10))

        self._kv(meta, "Enable:", self.enable_rb, 0)
        self._kv(meta, "Blade:", self.blade_rb, 1)
        self._kv(meta, "In Ref:", self.inref_rb, 2)
        self._kv(meta, "Out Ref:", self.outref_rb, 3)

        # Tabs
        nb = ttk.Notebook(right)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        tab_run = ttk.Frame(nb, style="Panel2.TFrame")
        tab_setup = ttk.Frame(nb, style="Panel2.TFrame")
        tab_display = ttk.Frame(nb, style="Panel2.TFrame")
        tab_console = ttk.Frame(nb, style="Panel2.TFrame")

        nb.add(tab_run, text="Run")
        nb.add(tab_setup, text="Setup")
        nb.add(tab_display, text="Display")
        nb.add(tab_console, text="Console")

        # RUN
        run_box = ttk.LabelFrame(tab_run, text="Run Control")
        run_box.pack(fill="x", padx=12, pady=12)

        ttk.Checkbutton(run_box, text="Enable (Run)", variable=self.enable_set, command=self.apply_enable).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(8, 10)
        )

        ttk.Label(run_box, text="Set Frequency (Hz):").grid(row=1, column=0, sticky="e", padx=10, pady=6)
        ttk.Entry(run_box, textvariable=self.freq_set, width=14).grid(row=1, column=1, sticky="w", pady=6)
        ttk.Button(run_box, text="Apply", command=self.apply_freq).grid(row=1, column=2, sticky="w", padx=8, pady=6)

        ref_box = ttk.LabelFrame(tab_run, text="Reference Routing")
        ref_box.pack(fill="x", padx=12, pady=(0, 12))

        ttk.Label(ref_box, text="Input Reference:").grid(row=0, column=0, sticky="e", padx=10, pady=6)
        self.inref_combo = ttk.Combobox(ref_box, state="disabled", width=22, textvariable=self.inref_choice)
        self.inref_combo.grid(row=0, column=1, sticky="w", pady=6)
        ttk.Button(ref_box, text="Apply", command=self.apply_inref).grid(row=0, column=2, sticky="w", padx=8, pady=6)

        ttk.Label(ref_box, text="Output Reference:").grid(row=1, column=0, sticky="e", padx=10, pady=6)
        self.outref_combo = ttk.Combobox(ref_box, state="disabled", width=22, textvariable=self.outref_choice)
        self.outref_combo.grid(row=1, column=1, sticky="w", pady=6)
        ttk.Button(ref_box, text="Apply", command=self.apply_outref).grid(row=1, column=2, sticky="w", padx=8, pady=6)

        # SETUP
        setup_box = ttk.LabelFrame(tab_setup, text="Blade / Harmonics / Phase")
        setup_box.pack(fill="x", padx=12, pady=12)

        ttk.Label(setup_box, text="Blade:").grid(row=0, column=0, sticky="e", padx=10, pady=6)
        self.blade_combo = ttk.Combobox(setup_box, state="disabled", width=22, textvariable=self.blade_choice,
                                        values=[b.name for b in Blade])
        self.blade_combo.grid(row=0, column=1, sticky="w", pady=6)
        ttk.Button(setup_box, text="Apply", command=self.apply_blade).grid(row=0, column=2, sticky="w", padx=8, pady=6)

        self._simple_set_row(setup_box, 1, "nharmonic:", self.nharm_set, lambda: self.apply_int_prop("nharmonic", self.nharm_set))
        self._simple_set_row(setup_box, 2, "dharmonic:", self.dharm_set, lambda: self.apply_int_prop("dharmonic", self.dharm_set))
        self._simple_set_row(setup_box, 3, "phase:", self.phase_set, lambda: self.apply_int_prop("phase", self.phase_set))
        self._simple_set_row(setup_box, 4, "oncycle:", self.oncycle_set, lambda: self.apply_int_prop("oncycle", self.oncycle_set))

        # DISPLAY
        disp_box = ttk.LabelFrame(tab_display, text="Display")
        disp_box.pack(fill="x", padx=12, pady=12)
        self._simple_set_row(disp_box, 0, "intensity:", self.intensity_set, lambda: self.apply_int_prop("intensity", self.intensity_set))

        # CONSOLE
        console_box = ttk.LabelFrame(tab_console, text="Console Log")
        console_box.pack(fill="both", expand=True, padx=12, pady=12)
        self.console = tk.Text(console_box, height=14, wrap="word",
                               bg="#ffffff", fg="#111111", insertbackground="#111111")
        self.console.pack(fill="both", expand=True, padx=10, pady=10)
        self._log("Ready. Type COM4 and click Connect.\n")

    def _simple_set_row(self, parent, row, label, var, cmd):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="e", padx=10, pady=6)
        ttk.Entry(parent, textvariable=var, width=14).grid(row=row, column=1, sticky="w", pady=6)
        ttk.Button(parent, text="Apply", command=cmd).grid(row=row, column=2, sticky="w", padx=8, pady=6)

    def _readout_card(self, parent, title, value_var, unit):
        card = ttk.Frame(parent, style="Panel2.TFrame")
        card.pack(fill="x", padx=10, pady=(10, 0))

        ttk.Label(card, text=title, style="Muted.TLabel", background=self.colors["panel2"]).pack(anchor="w", padx=10, pady=(10, 0))

        row = ttk.Frame(card, style="Panel2.TFrame")
        row.pack(fill="x", padx=10, pady=(0, 10))

        ttk.Label(row, textvariable=value_var, style="Big.TLabel", background=self.colors["panel2"]).pack(side="left")
        ttk.Label(row, text=unit, style="BigUnit.TLabel", background=self.colors["panel2"]).pack(side="left", padx=(10, 0))

    def _kv(self, parent, k, v_var, row):
        ttk.Label(parent, text=k).grid(row=row, column=0, sticky="w", padx=10, pady=4)
        ttk.Label(parent, textvariable=v_var, style="Muted.TLabel").grid(row=row, column=1, sticky="w", padx=10, pady=4)

    # ---------------- Status + logging ----------------
    def _set_status(self, text, kind="bad"):
        self.status_text.set(text)
        style = {"ok": "StatusOK.TLabel", "warn": "StatusWARN.TLabel", "bad": "StatusBAD.TLabel"}.get(kind, "StatusBAD.TLabel")
        self.status_label.configure(style=style)

    def _log(self, msg: str):
        self.console.insert("end", msg)
        self.console.see("end")

    # ---------------- Refresh loop ----------------
    def _start_auto_refresh(self):
        self._stop_auto_refresh()
        self._refresh_job = self.after(self.auto_refresh_ms, self._auto_refresh_tick)

    def _stop_auto_refresh(self):
        if self._refresh_job is not None:
            try:
                self.after_cancel(self._refresh_job)
            except Exception:
                pass
        self._refresh_job = None

    def _auto_refresh_tick(self):
        self.refresh_once(silent=True)
        self._refresh_job = self.after(self.auto_refresh_ms, self._auto_refresh_tick)

    # ---------------- Device actions ----------------
    def connect(self):
        if self.ch is not None:
            return
        port = self.port_var.get().strip()
        try:
            self.ch = MC2000B(serial_port=port)
            self.id_var.set(self.ch.id)
            self._set_status(f"Connected ({port})", kind="ok")

            self.blade_combo.configure(state="readonly")
            self.inref_combo.configure(state="readonly")
            self.outref_combo.configure(state="readonly")

            self._sync_controls_from_device()
            self._start_auto_refresh()

            self._log(f"[{datetime.now().strftime('%H:%M:%S')}] Connected. {self.ch.id}\n")
        except Exception as e:
            self.ch = None
            self.id_var.set("—")
            self._set_status("Disconnected", kind="bad")
            messagebox.showerror("Connect failed", str(e))
            self._log(f"ERROR connect: {e}\n")

    def disconnect(self):
        self._stop_auto_refresh()
        if self.ch is not None:
            try:
                self.ch.close()
            except Exception:
                pass
        self.ch = None
        self.id_var.set("—")
        self._set_status("Disconnected", kind="bad")

        for v in (self.freq_rb, self.input_rb, self.refout_rb, self.blade_rb, self.inref_rb, self.outref_rb, self.enable_rb):
            v.set("—")

        self.blade_combo.configure(state="disabled")
        self.inref_combo.configure(state="disabled")
        self.outref_combo.configure(state="disabled")

        self._log(f"[{datetime.now().strftime('%H:%M:%S')}] Disconnected.\n")

    def refresh_once(self, silent=False):
        if self.ch is None:
            return
        try:
            self.freq_rb.set(str(self.ch.freq))
            self.input_rb.set(str(self.ch.input))
            self.refout_rb.set(str(self.ch.refoutfreq))
            self.enable_rb.set("RUNNING" if int(self.ch.enable) == 1 else "STOPPED")
            self.blade_rb.set(self.ch.get_blade_string())
            self.inref_rb.set(self.ch.get_inref_string())
            self.outref_rb.set(self.ch.get_outref_string())

            self.enable_set.set(int(self.ch.enable))
            self._refresh_ref_lists()

            if not silent:
                self._log(f"[{datetime.now().strftime('%H:%M:%S')}] Refreshed.\n")
        except Exception as e:
            self._set_status("Connected (read error)", kind="warn")
            if not silent:
                messagebox.showerror("Refresh failed", str(e))
            self._log(f"ERROR refresh: {e}\n")

    def _sync_controls_from_device(self):
        # blade
        bname = self.ch.get_blade_string()
        self.blade_choice.set(bname)
        self.blade_combo.set(bname)

        # refresh inref/outref lists based on blade
        self._refresh_ref_lists(force=True)

        self.inref_choice.set(self.ch.get_inref_string())
        self.outref_choice.set(self.ch.get_outref_string())

        # seed other controls
        self.freq_set.set(str(self.ch.freq))
        self.nharm_set.set(str(self.ch.nharmonic))
        self.dharm_set.set(str(self.ch.dharmonic))
        self.phase_set.set(str(self.ch.phase))
        self.oncycle_set.set(str(self.ch.oncycle))
        self.intensity_set.set(str(self.ch.intensity))

    def _refresh_ref_lists(self, force=False):
        if self.ch is None:
            return
        try:
            blade = Blade(self.ch.blade)
            inrefs = list(blade.inrefs)
            outrefs = list(blade.outrefs)

            if force or tuple(self.inref_combo["values"]) != tuple(inrefs):
                self.inref_combo["values"] = inrefs
            if force or tuple(self.outref_combo["values"]) != tuple(outrefs):
                self.outref_combo["values"] = outrefs

            if self.inref_choice.get() not in inrefs:
                self.inref_choice.set(inrefs[0] if inrefs else "")
            if self.outref_choice.get() not in outrefs:
                self.outref_choice.set(outrefs[0] if outrefs else "")
        except Exception:
            pass

    # Apply handlers
    def apply_enable(self):
        if self.ch is None:
            return
        try:
            self.ch.enable = int(self.enable_set.get())
            self._log(f"Set enable={self.ch.enable}\n")
        except Exception as e:
            messagebox.showerror("Enable failed", str(e))
            self._log(f"ERROR enable: {e}\n")

    def apply_freq(self):
        if self.ch is None:
            return
        try:
            self.ch.freq = int(float(self.freq_set.get()))
            self._log(f"Set freq={self.ch.freq}\n")
        except Exception as e:
            messagebox.showerror("Set Frequency failed", str(e))
            self._log(f"ERROR freq: {e}\n")

    def apply_blade(self):
        if self.ch is None:
            return
        try:
            name = self.blade_choice.get().strip() or self.blade_combo.get().strip()
            self.ch.blade = int(Blade[name].value)
            self._log(f"Set blade={name}\n")
            self._refresh_ref_lists(force=True)
        except Exception as e:
            messagebox.showerror("Set Blade failed", str(e))
            self._log(f"ERROR blade: {e}\n")

    def apply_inref(self):
        if self.ch is None:
            return
        try:
            val = self.inref_choice.get().strip()
            self.ch.set_inref_string(val)
            self._log(f"Set input ref='{val}'\n")
        except Exception as e:
            messagebox.showerror("Set Input Reference failed", str(e))
            self._log(f"ERROR inref: {e}\n")

    def apply_outref(self):
        if self.ch is None:
            return
        try:
            val = self.outref_choice.get().strip()
            self.ch.set_outref_string(val)
            self._log(f"Set output ref='{val}'\n")
        except Exception as e:
            messagebox.showerror("Set Output Reference failed", str(e))
            self._log(f"ERROR outref: {e}\n")

    def apply_int_prop(self, prop: str, var: tk.StringVar):
        if self.ch is None:
            return
        try:
            val = int(float(var.get()))
            setattr(self.ch, prop, val)
            self._log(f"Set {prop}={getattr(self.ch, prop)}\n")
        except Exception as e:
            messagebox.showerror(f"Set {prop} failed", str(e))
            self._log(f"ERROR {prop}: {e}\n")

    def on_close(self):
        self.disconnect()
        self.destroy()


if __name__ == "__main__":
    app = MC2000BApp()
    app.mainloop()