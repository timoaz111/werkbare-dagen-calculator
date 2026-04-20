"""
Werkzaamheden Weekrapport — Tkinter GUI
Wekelijks bouwrapport met weersoverzicht, werkbaarheid, uren en werkzaamheden.
"""

import sys
import os
import threading
from datetime import date, timedelta

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from weekrapport_data import WeekRapportData, DAYS
from utils import ms_to_beaufort

DAYS_NL = {
    "ma": "Maandag", "di": "Dinsdag", "wo": "Woensdag",
    "do": "Donderdag", "vr": "Vrijdag", "za": "Zaterdag", "zo": "Zondag"
}
DAYS_SHORT = {
    "ma": "Ma", "di": "Di", "wo": "Wo",
    "do": "Do", "vr": "Vr", "za": "Za", "zo": "Zo"
}
WORKDAYS = ["ma", "di", "wo", "do", "vr"]

# Vaste pixelbreedtes werknemerstabel: Naam | Bedrijf | Functie | Opmerking | Ma..Zo | Totaal | 40u | 0u | X
WN_COL_WIDTHS = [150, 150, 110, 150, 42, 42, 42, 42, 42, 42, 42, 52, 40, 40, 28]
WN_ROW_H = 26

APP_BG = "#f8f9fa"
HEADER_BG = "#343a40"
HEADER_FG = "#ffffff"
PRIMARY = "#0d6efd"
DANGER = "#dc3545"
SUCCESS = "#198754"
SECTION_BG = "#ffffff"


def _wn_cell(parent, width, height=WN_ROW_H) -> tk.Frame:
    """Maakt een vaste-breedte cel-frame voor de werknemerstabel."""
    f = tk.Frame(parent, width=width, height=height, bg=APP_BG)
    f.pack_propagate(False)
    f.pack(side="left")
    return f


def bind_autocomplete(entry: ttk.Entry, get_values):
    """
    Voegt autocomplete toe aan een Entry: toont een popup-listbox met
    suggesties terwijl de gebruiker typt. Focus blijft in het invoerveld.
    """
    popup = {"win": None, "lb": None}

    def show(matches):
        close()
        if not matches:
            return
        win = tk.Toplevel(entry)
        win.wm_overrideredirect(True)
        win.wm_attributes("-topmost", True)

        lb = tk.Listbox(win, font=("Segoe UI", 9), selectbackground=PRIMARY,
                        selectforeground="white", activestyle="none",
                        height=min(6, len(matches)), relief="solid", bd=1)
        sb = ttk.Scrollbar(win, orient="vertical", command=lb.yview)
        lb.config(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True)
        if len(matches) > 6:
            sb.pack(side="right", fill="y")

        for m in matches:
            lb.insert("end", m)

        # Positie onder het invoerveld
        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        w = max(entry.winfo_width(), 160)
        row_h = 18
        win.geometry(f"{w}x{min(6, len(matches)) * row_h + 4}+{x}+{y}")

        def select(evt=None):
            sel = lb.curselection()
            if sel:
                entry.delete(0, "end")
                entry.insert(0, lb.get(sel[0]))
            close()
            entry.focus_set()

        lb.bind("<ButtonRelease-1>", select)
        lb.bind("<Return>", select)

        # Pijltje-omlaag vanuit entry → focus naar listbox
        def focus_lb(evt=None):
            lb.focus_set()
            if lb.size():
                lb.selection_set(0)
                lb.activate(0)
        entry.bind("<Down>", focus_lb)

        # Pijltje-omhoog vanuit listbox terug naar entry
        def back_to_entry(evt=None):
            if lb.curselection() and lb.curselection()[0] == 0:
                entry.focus_set()
        lb.bind("<Up>", back_to_entry)
        lb.bind("<Escape>", lambda e: (close(), entry.focus_set()))

        popup["win"] = win
        popup["lb"] = lb

    def close(evt=None):
        if popup["win"] and popup["win"].winfo_exists():
            popup["win"].destroy()
        popup["win"] = None
        popup["lb"] = None

    def on_key(evt):
        if evt.keysym in ("Tab", "Escape", "Return"):
            close()
            return
        if evt.keysym in ("Down", "Up", "Left", "Right"):
            return
        entry.after_idle(_update)

    def _update():
        tekst = entry.get().lower()
        alle = get_values()
        matches = [v for v in alle if v.lower().startswith(tekst)] if tekst else []
        if matches:
            show(matches)
        else:
            close()

    entry.bind("<KeyRelease>", on_key)
    entry.bind("<FocusOut>", lambda e: entry.after(150, close))


class WerknemerRij:
    def __init__(self, parent_frame, on_delete, on_change, get_bedrijven, get_namen, get_functies):
        self.on_delete = on_delete
        self.on_change = on_change
        self.get_bedrijven = get_bedrijven
        self.get_namen = get_namen
        self.get_functies = get_functies

        self.naam_var = tk.StringVar()
        self.bedrijf_var = tk.StringVar()
        self.functie_var = tk.StringVar()
        self.opmerking_var = tk.StringVar()
        self.uren_vars = {dag: tk.StringVar(value="0") for dag in DAYS}
        self.totaal_var = tk.StringVar(value="0")

        self.frame = tk.Frame(parent_frame, bg=APP_BG)
        self.frame.pack(fill="x", pady=1)

        widths = WN_COL_WIDTHS

        # Naam — Entry met autocomplete popup
        cell = _wn_cell(self.frame, widths[0])
        self._naam_entry = ttk.Entry(cell, textvariable=self.naam_var, font=("Segoe UI", 9))
        self._naam_entry.pack(fill="both", expand=True, padx=1, pady=1)
        bind_autocomplete(self._naam_entry, self.get_namen)

        # Bedrijf — Combobox (alleen dropdown)
        cell = _wn_cell(self.frame, widths[1])
        self._bedrijf_cb = ttk.Combobox(cell, textvariable=self.bedrijf_var, state="readonly", font=("Segoe UI", 9))
        self._bedrijf_cb.pack(fill="both", expand=True, padx=1, pady=1)
        self._bedrijf_cb.bind("<ButtonPress-1>", self._refresh_bedrijf_values)

        # Functie — Entry met autocomplete popup
        cell = _wn_cell(self.frame, widths[2])
        self._functie_entry = ttk.Entry(cell, textvariable=self.functie_var, font=("Segoe UI", 9))
        self._functie_entry.pack(fill="both", expand=True, padx=1, pady=1)
        bind_autocomplete(self._functie_entry, self.get_functies)

        # Opmerking — direct tekstveld
        cell = _wn_cell(self.frame, widths[3])
        ttk.Entry(cell, textvariable=self.opmerking_var, font=("Segoe UI", 9)).pack(
            fill="both", expand=True, padx=1, pady=1
        )

        # Uren per dag (7 kolommen)
        self._dag_cells: dict[str, tk.Frame] = {}
        for dag, w in zip(DAYS, widths[4:11]):
            cell = _wn_cell(self.frame, w)
            e = ttk.Entry(cell, textvariable=self.uren_vars[dag], justify="center")
            e.pack(fill="both", expand=True, padx=1, pady=1)
            self.uren_vars[dag].trace("w", self._recalc_totaal)
            self._dag_cells[dag] = cell

        # Totaal
        cell = _wn_cell(self.frame, widths[11])
        tk.Label(
            cell, textvariable=self.totaal_var,
            font=("Segoe UI", 9, "bold"), bg="#e9ecef", relief="groove", anchor="center"
        ).pack(fill="both", expand=True, padx=1, pady=1)

        # Fulltime-knop (40u: Ma-Vr = 8)
        cell = _wn_cell(self.frame, widths[12])
        tk.Button(
            cell, text="40u", fg="white", bg="#0d6efd",
            font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2",
            command=self._vul_fulltime
        ).pack(fill="both", expand=True, padx=1, pady=1)

        # Reset-knop (0u: alles naar 0)
        cell = _wn_cell(self.frame, widths[13])
        tk.Button(
            cell, text="0u", fg="white", bg="#6c757d",
            font=("Segoe UI", 8, "bold"), relief="flat", cursor="hand2",
            command=self._reset_uren
        ).pack(fill="both", expand=True, padx=1, pady=1)

        # Verwijder-knop
        cell = _wn_cell(self.frame, widths[14])
        tk.Button(
            cell, text="✕", fg=DANGER, bg=APP_BG,
            font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
            command=self._delete
        ).pack(fill="both", expand=True)

        # Trace voor auto-opslaan
        for var in [self.naam_var, self.bedrijf_var, self.functie_var, self.opmerking_var]:
            var.trace("w", lambda *_: on_change())

    def set_za_zo_visible(self, zichtbaar: bool):
        for dag in ["za", "zo"]:
            cell = self._dag_cells.get(dag)
            if cell:
                if zichtbaar:
                    cell.pack(side="left")
                else:
                    cell.pack_forget()

    def _vul_fulltime(self):
        for dag in WORKDAYS:
            self.uren_vars[dag].set("8")
        for dag in ["za", "zo"]:
            self.uren_vars[dag].set("0")

    def _reset_uren(self):
        for dag in DAYS:
            self.uren_vars[dag].set("0")

    def _refresh_bedrijf_values(self, *_):
        self._bedrijf_cb["values"] = self.get_bedrijven()

    def _recalc_totaal(self, *_):
        totaal = 0
        for dag in DAYS:
            try:
                totaal += float(self.uren_vars[dag].get() or 0)
            except ValueError:
                pass
        self.totaal_var.set(str(int(totaal)) if totaal == int(totaal) else str(round(totaal, 1)))
        self.on_change()

    def _delete(self):
        self.frame.destroy()
        self.on_delete(self)

    def get_data(self) -> dict:
        uren = {}
        for dag in DAYS:
            try:
                uren[dag] = float(self.uren_vars[dag].get() or 0)
            except ValueError:
                uren[dag] = 0
        return {
            "naam": self.naam_var.get(),
            "bedrijf": self.bedrijf_var.get(),
            "functie": self.functie_var.get(),
            "uren": uren,
            "opmerking": self.opmerking_var.get(),
        }

    def set_data(self, data: dict):
        self.naam_var.set(data.get("naam", ""))
        self.bedrijf_var.set(data.get("bedrijf", ""))
        self.functie_var.set(data.get("functie", ""))
        uren = data.get("uren", {})
        for dag in DAYS:
            val = uren.get(dag, 0)
            self.uren_vars[dag].set(str(int(val)) if val == int(val) else str(val))
        self.opmerking_var.set(data.get("opmerking", ""))


class DagWerkzaamhedenPanel:
    def __init__(self, parent, dag: str, on_change):
        self.dag = dag
        self.on_change = on_change
        self.rijen: list[tuple[tk.Frame, tk.StringVar]] = []

        self.frame = tk.LabelFrame(
            parent, text=DAYS_NL.get(dag, dag),
            font=("Segoe UI", 9, "bold"), bg=APP_BG, padx=8, pady=6
        )
        self.frame.pack(fill="x", pady=4)

        self.rows_frame = tk.Frame(self.frame, bg=APP_BG)
        self.rows_frame.pack(fill="x")

        btn_frame = tk.Frame(self.frame, bg=APP_BG)
        btn_frame.pack(fill="x", pady=(4, 0))
        tk.Button(
            btn_frame, text="+ Activiteit toevoegen",
            font=("Segoe UI", 9), bg=APP_BG, fg=PRIMARY,
            relief="flat", cursor="hand2",
            command=lambda: self._add_rij()
        ).pack(side="left")

    def _add_rij(self, tekst: str = "") -> tk.StringVar:
        var = tk.StringVar(value=tekst)
        row_frame = tk.Frame(self.rows_frame, bg=APP_BG)
        row_frame.pack(fill="x", pady=2)

        entry = ttk.Entry(row_frame, textvariable=var, width=70)
        entry.pack(side="left", padx=(0, 6), fill="x", expand=True)

        def delete_this():
            row_frame.destroy()
            self.rijen = [(f, v) for f, v in self.rijen if f != row_frame]
            self.on_change()

        tk.Button(
            row_frame, text="✕", fg=DANGER, bg=APP_BG,
            font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2",
            command=delete_this
        ).pack(side="left")

        self.rijen.append((row_frame, var))
        var.trace("w", lambda *_: self.on_change())
        return var

    def set_activiteiten(self, activiteiten: list[str]):
        # Verwijder bestaande rijen
        for row_frame, _ in self.rijen:
            row_frame.destroy()
        self.rijen = []
        # Voeg nieuwe toe (minimaal 3)
        items = activiteiten if activiteiten else ["", "", ""]
        if len(items) < 3:
            items = items + [""] * (3 - len(items))
        for tekst in items:
            self._add_rij(tekst)

    def get_activiteiten(self) -> list[str]:
        return [v.get() for _, v in self.rijen]


class WeekRapportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Werkzaamheden Weekrapport")
        self.geometry("1250x900")
        self.minsize(1000, 700)
        self.configure(bg=APP_BG)

        try:
            if getattr(sys, 'frozen', False):
                icon_pad = os.path.join(sys._MEIPASS, "tools", "icon.ico")
            else:
                icon_pad = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "icon.ico")
            self.iconbitmap(icon_pad)
        except Exception:
            pass

        # Staat
        today = date.today()
        iso = today.isocalendar()
        self.current_iso_year = iso[0]
        self.current_kalender_week = iso[1]
        self._export_start_iso_year = iso[0]
        self._export_start_week = iso[1]
        self._export_end_iso_year = iso[0]
        self._export_end_week = iso[1]
        self.project_start_var = tk.StringVar()
        self.project_start_var.trace("w", self._on_project_start_changed)
        self.project_end_var = tk.StringVar()
        self.project_end_var.trace("w", self._on_project_end_changed)
        self._projecten: list[str] = WeekRapportData.load_projecten()
        initial_project = self._projecten[0] if self._projecten else "P000000"
        self.project_nr_var = tk.StringVar(value=initial_project)
        self.project_nr_var.trace("w", self._on_change)
        self._werknemer_rijen: list[WerknemerRij] = []
        self._dag_panels: dict[str, DagWerkzaamhedenPanel] = {}
        self._autosave_after = None
        self._loading = False  # voorkomt autosave tijdens laden
        self._bedrijven: list[str] = WeekRapportData.load_bedrijven()
        self._instellingen: dict = WeekRapportData.load_instellingen()
        self._werknemers_namen: list[str] = WeekRapportData.get_all_unique_namen()
        self._werknemers_functies: list[str] = WeekRapportData.get_all_unique_functies()
        self._za_zo_zichtbaar = False

        self._build_ui()
        self._build_menu()
        self._load_week()
        self._apply_instellingen()

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_menu(self):
        menubar = tk.Menu(self)
        extra_menu = tk.Menu(menubar, tearoff=0)
        extra_menu.add_command(label="Nieuw project", command=self._nieuw_project)
        extra_menu.add_command(label="Ander project openen", command=self._open_project_kiezen)
        extra_menu.add_separator()
        extra_menu.add_command(label="Calculator openen", command=self._open_calculator)
        extra_menu.add_separator()
        extra_menu.add_command(label="Instellingen", command=self._open_instellingen)
        menubar.add_cascade(label="Extra", menu=extra_menu)
        self.config(menu=menubar)

    def _open_calculator(self):
        from gui import WerkbareDagenApp
        WerkbareDagenApp(self)

    def _apply_instellingen(self):
        weer_verbergen = self._instellingen.get("weer_za_zo_verbergen", False)
        wn_verbergen   = self._instellingen.get("werknemers_za_zo_verbergen", False)

        # Weer rapport: za/zo rijen tonen of verbergen
        for dag in ["za", "zo"]:
            for w in self._weer_dag_widgets.get(dag, []):
                if weer_verbergen:
                    w.grid_remove()
                else:
                    w.grid()

        # Werknemers: header cellen
        for f in self._wn_header_za_zo:
            if wn_verbergen:
                f.pack_forget()
            else:
                f.pack(side="left")

        # Werknemers: elke rij
        for rij in self._werknemer_rijen:
            rij.set_za_zo_visible(not wn_verbergen)

    def _open_instellingen(self):
        dlg = tk.Toplevel(self)
        dlg.title("Instellingen")
        dlg.resizable(False, False)
        dlg.grab_set()

        pad = {"padx": 16, "pady": 8}

        # ── Instellingen Weer Rapport ──────────────────────────
        grp1 = tk.LabelFrame(dlg, text="Instellingen Weer Rapport",
                             font=("Segoe UI", 9, "bold"), padx=12, pady=8)
        grp1.pack(fill="x", padx=16, pady=(12, 4))
        weer_var = tk.BooleanVar(value=self._instellingen.get("weer_za_zo_verbergen", False))
        tk.Checkbutton(grp1, text="Zaterdag en Zondag verbergen",
                       variable=weer_var, font=("Segoe UI", 9)).pack(anchor="w")

        # ── Instellingen Uren Werknemers ───────────────────────
        grp2 = tk.LabelFrame(dlg, text="Instellingen Uren Werknemers",
                             font=("Segoe UI", 9, "bold"), padx=12, pady=8)
        grp2.pack(fill="x", padx=16, pady=(4, 12))
        wn_var = tk.BooleanVar(value=self._instellingen.get("werknemers_za_zo_verbergen", False))
        tk.Checkbutton(grp2, text="Zaterdag en Zondag verbergen",
                       variable=wn_var, font=("Segoe UI", 9)).pack(anchor="w")

        # ── Knoppen ────────────────────────────────────────────
        btn_row = tk.Frame(dlg)
        btn_row.pack(pady=(0, 12))

        def opslaan():
            self._instellingen["weer_za_zo_verbergen"] = weer_var.get()
            self._instellingen["werknemers_za_zo_verbergen"] = wn_var.get()
            WeekRapportData.save_instellingen(self._instellingen)
            self._apply_instellingen()
            dlg.destroy()

        tk.Button(btn_row, text="Opslaan", command=opslaan,
                  font=("Segoe UI", 9, "bold"), bg=PRIMARY, fg="white",
                  relief="flat", padx=16, pady=4, cursor="hand2").pack(side="left", padx=(0, 8))
        tk.Button(btn_row, text="Annuleren", command=dlg.destroy,
                  font=("Segoe UI", 9), bg="#f8f9fa", relief="flat",
                  padx=12, pady=4, cursor="hand2").pack(side="left")

    # ──────────────────────────────────────────── UI opbouwen

    def _build_ui(self):
        # ttk stijlen voor project-start invoerveld
        style = ttk.Style(self)
        style.configure("Valid.TEntry", fieldbackground="#d4edda")
        style.configure("Invalid.TEntry", fieldbackground="#f8d7da")
        style.configure("ValidEnd.TEntry", fieldbackground="#cce5ff")
        style.configure("InvalidEnd.TEntry", fieldbackground="#f8d7da")

        # Vaste header balk (buiten scrollgebied)
        self._build_header_bar()

        # Scrollbaar gedeelte
        outer = tk.Frame(self, bg=APP_BG)
        outer.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(outer, bg=APP_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._canvas, bg=APP_BG)
        self._canvas_window = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")

        self._inner.bind(
            "<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all"))
        )
        self._canvas.bind(
            "<Configure>",
            lambda e: self._canvas.itemconfig(self._canvas_window, width=e.width)
        )
        self._canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # Secties in inner frame
        self._build_weer_section()
        self._build_werkbaarheid_section()
        self._build_werknemers_section()
        self._build_werkzaamheden_section()

    def _on_mousewheel(self, event):
        self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _build_header_bar(self):
        bar = tk.Frame(self, bg=HEADER_BG, pady=10)
        bar.pack(fill="x")

        left = tk.Frame(bar, bg=HEADER_BG)
        left.pack(side="left", padx=16)

        tk.Label(left, text="Project nm:", font=("Segoe UI", 9), bg=HEADER_BG, fg="#adb5bd").pack(side="left")
        ttk.Entry(left, textvariable=self.project_nr_var, width=14,
                  font=("Segoe UI", 10)).pack(side="left", padx=(4, 16))

        tk.Label(left, text="Weeknr:", font=("Segoe UI", 9), bg=HEADER_BG, fg="#adb5bd").pack(side="left")
        self.project_week_nr_var = tk.StringVar(value="1")
        ttk.Entry(left, textvariable=self.project_week_nr_var, width=10,
                  font=("Segoe UI", 10)).pack(side="left", padx=(4, 16))
        self.project_week_nr_var.trace("w", self._on_change)

        tk.Label(left, text="Project start:", font=("Segoe UI", 9), bg=HEADER_BG, fg="#adb5bd").pack(side="left")
        self._project_start_entry = ttk.Entry(left, textvariable=self.project_start_var,
                                              width=12, font=("Segoe UI", 10))
        self._project_start_entry.pack(side="left", padx=(4, 4))
        tk.Label(left, text="dd-mm-jjjj", font=("Segoe UI", 8), bg=HEADER_BG, fg="#6c757d").pack(side="left", padx=(0, 12))

        tk.Label(left, text="Project einde:", font=("Segoe UI", 9), bg=HEADER_BG, fg="#adb5bd").pack(side="left")
        self._project_end_entry = ttk.Entry(left, textvariable=self.project_end_var,
                                            width=12, font=("Segoe UI", 10))
        self._project_end_entry.pack(side="left", padx=(4, 4))
        tk.Label(left, text="dd-mm-jjjj", font=("Segoe UI", 8), bg=HEADER_BG, fg="#6c757d").pack(side="left", padx=(0, 16))

        self._week_label = tk.Label(
            left, text="", font=("Segoe UI", 10, "bold"),
            bg=HEADER_BG, fg=HEADER_FG
        )
        self._week_label.pack(side="left", padx=(0, 16))

        self._datum_label = tk.Label(
            left, text="", font=("Segoe UI", 9),
            bg=HEADER_BG, fg="#adb5bd"
        )
        self._datum_label.pack(side="left")

        right = tk.Frame(bar, bg=HEADER_BG)
        right.pack(side="right", padx=16)

        tk.Button(
            right, text="< Vorige week", command=lambda: self._nav_week(-1),
            font=("Segoe UI", 9), bg="#495057", fg=HEADER_FG,
            relief="flat", padx=10, pady=4, cursor="hand2",
            activebackground="#6c757d"
        ).pack(side="left", padx=(0, 6))

        tk.Button(
            right, text="Volgende week >", command=lambda: self._nav_week(1),
            font=("Segoe UI", 9), bg="#495057", fg=HEADER_FG,
            relief="flat", padx=10, pady=4, cursor="hand2",
            activebackground="#6c757d"
        ).pack(side="left", padx=(0, 24))

        tk.Button(
            right, text="Export Excel",
            command=self._export_excel,
            font=("Segoe UI", 9, "bold"), bg=SUCCESS, fg="white",
            relief="flat", padx=14, pady=6, cursor="hand2",
            activebackground="#157347"
        ).pack(side="left")

        self._status_label = tk.Label(
            bar, text="", font=("Segoe UI", 8, "italic"),
            bg=HEADER_BG, fg="#6c757d"
        )
        self._status_label.pack(side="left", padx=16)

    def _build_weer_section(self):
        outer = tk.LabelFrame(
            self._inner, text="Weer Rapport",
            font=("Segoe UI", 10, "bold"), bg=APP_BG, padx=12, pady=10
        )
        outer.pack(fill="x", padx=16, pady=(12, 0))

        # Kolomkoppen
        headers = ["Dag", "Weer", "Temperatuur (°C)", "Regen (mm)", "Windkracht (Bft)", "Onwerkbaar", "Opmerking"]
        col_widths = [10, 24, 16, 12, 16, 12, 28]
        for c, (h, w) in enumerate(zip(headers, col_widths)):
            tk.Label(
                outer, text=h, font=("Segoe UI", 9, "bold"),
                bg="#343a40", fg="white", width=w, anchor="center", pady=4
            ).grid(row=0, column=c, padx=2, pady=(0, 4), sticky="ew")

        self._weer_vars: dict[str, dict[str, tk.StringVar]] = {}
        self._weer_dag_widgets: dict[str, list] = {}  # dag → lijst van widgets op die rij
        self._handmatig_onwerkbaar_vars: dict[str, tk.BooleanVar] = {}
        self._handmatig_reden_vars: dict[str, tk.StringVar] = {}
        self._handmatig_onwerkbaar_btns: dict[str, tk.Button] = {}
        for r, dag in enumerate(DAYS, 1):
            self._weer_vars[dag] = {
                "beschrijving": tk.StringVar(),
                "temp_c": tk.StringVar(),
                "regen_mm": tk.StringVar(),
                "wind_bft": tk.StringVar(),
            }

            bg = "#ffffff" if r % 2 == 0 else "#f8f9fa"
            rij_widgets = []

            lbl = tk.Label(
                outer, text=DAYS_NL[dag],
                font=("Segoe UI", 9, "bold"), bg=bg, anchor="center", width=10, pady=3
            )
            lbl.grid(row=r, column=0, padx=2, pady=1, sticky="ew")
            rij_widgets.append(lbl)

            for c, key in enumerate(["beschrijving", "temp_c", "regen_mm", "wind_bft"], 1):
                w = col_widths[c]
                e = ttk.Entry(
                    outer, textvariable=self._weer_vars[dag][key],
                    width=w, justify="center" if c > 1 else "left"
                )
                e.grid(row=r, column=c, padx=2, pady=1, sticky="ew")
                self._weer_vars[dag][key].trace("w", lambda *_: self._on_change())
                rij_widgets.append(e)

            # Handmatig onwerkbaar (alleen werkdagen tellen mee)
            self._handmatig_onwerkbaar_vars[dag] = tk.BooleanVar(value=False)
            self._handmatig_reden_vars[dag] = tk.StringVar(value="")
            if dag in WORKDAYS:
                btn = tk.Button(
                    outer, text="Onwerkbaar",
                    font=("Segoe UI", 8), relief="groove", cursor="hand2",
                    bg="#e9ecef", fg="#495057", pady=1,
                    command=lambda d=dag: self._toggle_handmatig_onwerkbaar(d)
                )
                btn.grid(row=r, column=5, padx=2, pady=1, sticky="ew")
                self._handmatig_onwerkbaar_btns[dag] = btn
                rij_widgets.append(btn)
            else:
                lbl = tk.Label(outer, text="", bg=bg, width=12)
                lbl.grid(row=r, column=5, padx=2, pady=1, sticky="ew")
                rij_widgets.append(lbl)

            # Opmerking vrij tekstveld (alle dagen)
            e_opm = ttk.Entry(outer, textvariable=self._handmatig_reden_vars[dag],
                              width=col_widths[6], justify="left")
            e_opm.grid(row=r, column=6, padx=2, pady=1, sticky="ew")
            self._handmatig_reden_vars[dag].trace("w", lambda *_: self._on_change())
            rij_widgets.append(e_opm)

            self._weer_dag_widgets[dag] = rij_widgets

        # Status label voor weer-laden
        self._weer_status = tk.Label(
            outer, text="", font=("Segoe UI", 8, "italic"),
            bg=APP_BG, fg="#6c757d"
        )
        self._weer_status.grid(row=len(DAYS) + 1, column=0, columnspan=7, sticky="w", pady=(4, 0))

    def _build_werkbaarheid_section(self):
        outer = tk.LabelFrame(
            self._inner, text="Werkbaarheid",
            font=("Segoe UI", 10, "bold"), bg=APP_BG, padx=12, pady=10
        )
        outer.pack(fill="x", padx=16, pady=(10, 0))

        self._wb_vars = {
            "werkbare_dagen": tk.StringVar(value="—"),
            "onwerkbare_dagen": tk.StringVar(value="—"),
            "onwerkbaar_feestdagen": tk.StringVar(value="—"),
            "onwerkbaar_weer": tk.StringVar(value="—"),
        }
        items = [
            ("Werkbare dagen (Ma–Vr):", "werkbare_dagen", "#d4edda"),
            ("Onwerkbare dagen (Ma–Vr):", "onwerkbare_dagen", "#f8d7da"),
            ("  w.v. Feestdagen:", "onwerkbaar_feestdagen", "#fff3cd"),
            ("  w.v. Weersomstandigheden:", "onwerkbaar_weer", "#ffe8b0"),
        ]
        for c, (label, key, kleur) in enumerate(items):
            tk.Label(
                outer, text=label, font=("Segoe UI", 9, "bold"),
                bg=APP_BG, anchor="w"
            ).grid(row=0, column=c * 2, sticky="w", padx=(0, 6), pady=4)
            tk.Label(
                outer, textvariable=self._wb_vars[key],
                font=("Segoe UI", 12, "bold"), bg=kleur,
                width=5, anchor="center", relief="groove"
            ).grid(row=0, column=c * 2 + 1, padx=(0, 24), pady=4)

        self._feestdag_title = tk.Label(
            outer, text="Feestdag:", font=("Segoe UI", 9, "bold"),
            bg=APP_BG, fg="#212529", anchor="w"
        )
        self._feestdag_title.grid(row=1, column=0, sticky="w", pady=(4, 0))
        self._feestdag_title.grid_remove()

        self._feestdagen_label = tk.Label(
            outer, text="", font=("Segoe UI", 9),
            bg=APP_BG, fg="#212529", anchor="w"
        )
        self._feestdagen_label.grid(row=1, column=1, columnspan=7, sticky="w", pady=(4, 0))
        self._feestdagen_label.grid_remove()  # verbergen als er geen feestdagen zijn

        self._wb_status = tk.Label(
            outer, text="", font=("Segoe UI", 8, "italic"),
            bg=APP_BG, fg="#6c757d"
        )
        self._wb_status.grid(row=2, column=0, columnspan=8, sticky="w")

    def _build_werknemers_section(self):
        outer = tk.LabelFrame(
            self._inner, text="Uren Werknemers",
            font=("Segoe UI", 10, "bold"), bg=APP_BG, padx=12, pady=10
        )
        outer.pack(fill="x", padx=16, pady=(10, 0))

        # Knoppen boven de tabel
        top_row = tk.Frame(outer, bg=APP_BG)
        top_row.pack(fill="x", pady=(0, 6))
        tk.Button(
            top_row, text="⚙ Bedrijven beheren",
            font=("Segoe UI", 9, "bold"), bg="#495057", fg="white",
            relief="flat", cursor="hand2", padx=10, pady=3,
            activebackground="#343a40",
            command=self._open_bedrijven_beheer
        ).pack(side="left", padx=(0, 10))
        tk.Button(
            top_row, text="↩ Overnemen van vorige week",
            font=("Segoe UI", 9, "bold"), bg="#198754", fg="white",
            relief="flat", cursor="hand2", padx=10, pady=3,
            activebackground="#157347",
            command=self._overnemen_vorige_week
        ).pack(side="left", padx=(0, 10))
        tk.Button(
            top_row, text="↩ Overnemen incl. uren",
            font=("Segoe UI", 9, "bold"), bg="#0d6efd", fg="white",
            relief="flat", cursor="hand2", padx=10, pady=3,
            activebackground="#0b5ed7",
            command=lambda: self._overnemen_vorige_week(met_uren=True)
        ).pack(side="left")

        # Bedrijfsfilter
        filter_row = tk.Frame(outer, bg=APP_BG)
        filter_row.pack(fill="x", pady=(0, 6))
        tk.Label(filter_row, text="Filter op bedrijf:", font=("Segoe UI", 9),
                 bg=APP_BG).pack(side="left", padx=(0, 6))
        self._bedrijf_filter_var = tk.StringVar(value="Alle bedrijven")
        self._bedrijf_filter_cb = ttk.Combobox(
            filter_row, textvariable=self._bedrijf_filter_var,
            state="readonly", width=28, font=("Segoe UI", 9)
        )
        self._bedrijf_filter_cb.pack(side="left")
        self._bedrijf_filter_cb.bind("<<ComboboxSelected>>", lambda _: self._apply_bedrijf_filter())
        self._refresh_bedrijf_filter_values()

        # Kolomkoppen — zelfde pixelbreedtes als WerknemerRij
        headers = (
            ["Naam Werknemer", "Naam Bedrijf", "Functie", "Opmerking"] +
            [f"{DAYS_SHORT[d]}" for d in DAYS] +
            ["Totaal", "40u", "0u", ""]
        )
        header_row = tk.Frame(outer, bg="#343a40")
        header_row.pack(fill="x", pady=(0, 2))
        self._wn_header_za_zo = []
        za_zo_indices = {DAYS.index("za") + 4, DAYS.index("zo") + 4}  # +4 voor Naam/Bedrijf/Functie/Opmerking
        for i, (h, w) in enumerate(zip(headers, WN_COL_WIDTHS)):
            f = tk.Frame(header_row, width=w, height=WN_ROW_H, bg="#343a40")
            f.pack_propagate(False)
            f.pack(side="left")
            tk.Label(
                f, text=h, font=("Segoe UI", 9, "bold"),
                bg="#343a40", fg="white", anchor="center"
            ).pack(fill="both", expand=True)
            if i in za_zo_indices:
                self._wn_header_za_zo.append(f)

        # Container voor dynamische rijen
        self._werknemers_frame = tk.Frame(outer, bg=APP_BG)
        self._werknemers_frame.pack(fill="x")

        # Knoppen
        btn_row = tk.Frame(outer, bg=APP_BG)
        btn_row.pack(fill="x", pady=(8, 0))

        tk.Button(
            btn_row, text="+ Werknemer toevoegen",
            font=("Segoe UI", 9), bg=APP_BG, fg=PRIMARY,
            relief="flat", cursor="hand2",
            command=lambda: self._add_werknemer_rij()
        ).pack(side="left", padx=(0, 12))

        tk.Button(
            btn_row, text="Verwijder alle werknemers",
            font=("Segoe UI", 9), bg=APP_BG, fg=DANGER,
            relief="flat", cursor="hand2",
            command=self._clear_all_werknemers
        ).pack(side="left", padx=(0, 24))

    def _build_werkzaamheden_section(self):
        outer = tk.LabelFrame(
            self._inner, text="Werkzaamheden",
            font=("Segoe UI", 10, "bold"), bg=APP_BG, padx=12, pady=10
        )
        outer.pack(fill="x", padx=16, pady=(10, 16))

        # Weekdagen panels
        for dag in WORKDAYS:
            panel = DagWerkzaamhedenPanel(outer, dag, self._on_change)
            self._dag_panels[dag] = panel

        # Za/Zo toggle knop
        self._za_zo_btn = tk.Button(
            outer, text="+ Toon Za / Zo werkzaamheden",
            font=("Segoe UI", 9), bg=APP_BG, fg="#6c757d",
            relief="flat", cursor="hand2",
            command=self._toggle_za_zo
        )
        self._za_zo_btn.pack(anchor="w", pady=(4, 0))

        # Za en Zo panels (standaard verborgen)
        for dag in ["za", "zo"]:
            panel = DagWerkzaamhedenPanel(outer, dag, self._on_change)
            self._dag_panels[dag] = panel
            panel.frame.pack_forget()  # verbergen

    # ──────────────────────────────────────────── Werknemers

    def _add_werknemer_rij(self, naam="", bedrijf="", functie="",
                           uren: dict | None = None, opmerking: str = "") -> WerknemerRij:
        rij = WerknemerRij(
            self._werknemers_frame,
            on_delete=self._on_werknemer_delete,
            on_change=self._on_change,
            get_bedrijven=lambda: self._bedrijven,
            get_namen=lambda: self._werknemers_namen,
            get_functies=lambda: self._werknemers_functies,
        )
        if naam or bedrijf or functie or uren or opmerking:
            rij.set_data({
                "naam": naam, "bedrijf": bedrijf, "functie": functie,
                "uren": uren or {dag: 0 for dag in DAYS},
                "opmerking": opmerking,
            })
        self._werknemer_rijen.append(rij)
        if self._instellingen.get("werknemers_za_zo_verbergen", False):
            rij.set_za_zo_visible(False)
        return rij

    def _on_werknemer_delete(self, rij: WerknemerRij):
        if rij in self._werknemer_rijen:
            self._werknemer_rijen.remove(rij)
        self._on_change()

    def _overnemen_vorige_week(self, met_uren: bool = False):
        project_nr = self.project_nr_var.get().strip() or "P000000"
        vorige_maandag = (
            WeekRapportData.week_dates_from_iso(self.current_iso_year, self.current_kalender_week)[0]
            - timedelta(weeks=1)
        )
        iso = vorige_maandag.isocalendar()
        vorige_data = WeekRapportData.load(project_nr, iso[0], iso[1])
        if not vorige_data:
            messagebox.showinfo(
                "Geen data",
                f"Geen opgeslagen gegevens gevonden voor de vorige week ({vorige_maandag.strftime('%d-%m-%Y')}).",
                parent=self
            )
            return
        vorige_werknemers = vorige_data.get("werknemers", [])
        if not vorige_werknemers:
            messagebox.showinfo("Leeg", "Vorige week heeft geen werknemers.", parent=self)
            return
        for rij in list(self._werknemer_rijen):
            rij.frame.destroy()
        self._werknemer_rijen.clear()
        for wn in vorige_werknemers:
            self._add_werknemer_rij(
                naam=wn.get("naam", ""),
                bedrijf=wn.get("bedrijf", ""),
                functie=wn.get("functie", ""),
                uren=wn.get("uren") if met_uren else None
            )
        self._on_change()

    def _refresh_bedrijf_filter_values(self):
        opties = ["Alle bedrijven"] + list(self._bedrijven)
        self._bedrijf_filter_cb["values"] = opties

    def _apply_bedrijf_filter(self):
        keuze = self._bedrijf_filter_var.get()
        for rij in self._werknemer_rijen:
            if keuze == "Alle bedrijven" or rij.bedrijf_var.get() == keuze:
                rij.frame.pack(fill="x", pady=1)
            else:
                rij.frame.pack_forget()

    def _open_bedrijven_beheer(self):
        dlg = tk.Toplevel(self)
        dlg.title("Bedrijven beheren")
        dlg.resizable(False, False)
        dlg.grab_set()

        tk.Label(dlg, text="Bedrijven", font=("Segoe UI", 10, "bold"), padx=12, pady=8).pack(anchor="w")

        frame = tk.Frame(dlg, padx=12)
        frame.pack(fill="both", expand=True)

        listbox = tk.Listbox(frame, width=40, height=12, font=("Segoe UI", 9), selectmode="single")
        listbox.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
        sb.pack(side="left", fill="y")
        listbox.config(yscrollcommand=sb.set)

        for b in self._bedrijven:
            listbox.insert("end", b)

        invoer_frame = tk.Frame(dlg, padx=12, pady=6)
        invoer_frame.pack(fill="x")

        nieuw_var = tk.StringVar()
        ttk.Entry(invoer_frame, textvariable=nieuw_var, width=30,
                  font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))

        def toevoegen():
            naam = nieuw_var.get().strip()
            if naam and naam not in self._bedrijven:
                self._bedrijven.append(naam)
                self._bedrijven.sort()
                listbox.delete(0, "end")
                for b in self._bedrijven:
                    listbox.insert("end", b)
                WeekRapportData.save_bedrijven(self._bedrijven)
            nieuw_var.set("")

        def verwijderen():
            sel = listbox.curselection()
            if not sel:
                return
            naam = listbox.get(sel[0])
            self._bedrijven = [b for b in self._bedrijven if b != naam]
            listbox.delete(sel[0])
            WeekRapportData.save_bedrijven(self._bedrijven)

        tk.Button(invoer_frame, text="Toevoegen", command=toevoegen,
                  font=("Segoe UI", 9), bg=PRIMARY, fg="white",
                  relief="flat", padx=8, cursor="hand2").pack(side="left", padx=(0, 6))
        tk.Button(invoer_frame, text="Verwijder geselecteerde", command=verwijderen,
                  font=("Segoe UI", 9), bg=APP_BG, fg=DANGER,
                  relief="flat", padx=8, cursor="hand2").pack(side="left")

        nieuw_var_entry = invoer_frame.winfo_children()[0]
        nieuw_var_entry.bind("<Return>", lambda _: toevoegen())

        def sluiten():
            self._refresh_bedrijf_filter_values()
            dlg.destroy()

        tk.Button(dlg, text="Sluiten", command=sluiten,
                  font=("Segoe UI", 9), bg=APP_BG, relief="flat",
                  padx=12, pady=4, cursor="hand2").pack(pady=(0, 10))

    def _clear_all_werknemers(self):
        if not self._werknemer_rijen:
            return
        if not messagebox.askyesno(
            "Bevestigen",
            "Wil je alle werknemers verwijderen?",
            parent=self
        ):
            return
        for rij in list(self._werknemer_rijen):
            rij.frame.destroy()
        self._werknemer_rijen.clear()
        self._on_change()

    # ──────────────────────────────────────────── Za/Zo toggle

    def _toggle_za_zo(self):
        self._za_zo_zichtbaar = not self._za_zo_zichtbaar
        for dag in ["za", "zo"]:
            if self._za_zo_zichtbaar:
                self._dag_panels[dag].frame.pack(fill="x", pady=4)
            else:
                self._dag_panels[dag].frame.pack_forget()
        self._za_zo_btn.config(
            text="− Verberg Za / Zo werkzaamheden" if self._za_zo_zichtbaar
            else "+ Toon Za / Zo werkzaamheden"
        )

    # ──────────────────────────────────────────── Week navigatie

    def _get_project_start_monday(self):
        """Geeft de maandag van de project-startweek, of None als geen geldige datum ingevuld."""
        raw = self.project_start_var.get().strip()
        if not raw:
            return None
        try:
            from datetime import datetime
            d = datetime.strptime(raw, "%d-%m-%Y").date()
            # Maandag van die week
            return d - timedelta(days=d.weekday())
        except ValueError:
            return None

    def _on_project_start_changed(self, *_):
        raw = self.project_start_var.get().strip()
        if not raw:
            self._project_start_entry.configure(style="TEntry")
            return
        try:
            from datetime import datetime
            d = datetime.strptime(raw, "%d-%m-%Y").date()
            self._project_start_entry.configure(style="Valid.TEntry")
            maandag = d - timedelta(days=d.weekday())
            iso = maandag.isocalendar()
            self._export_start_iso_year = iso[0]
            self._export_start_week = iso[1]
            if hasattr(self, '_export_start_label'):
                self._update_export_start_label()
            self._propagate_project_dates()
        except ValueError:
            self._project_start_entry.configure(style="Invalid.TEntry")
        self._on_change()

    def _get_project_end_monday(self):
        """Geeft de maandag van de project-eindweek, of None als geen geldige datum ingevuld."""
        raw = self.project_end_var.get().strip()
        if not raw:
            return None
        try:
            from datetime import datetime
            d = datetime.strptime(raw, "%d-%m-%Y").date()
            return d - timedelta(days=d.weekday())
        except ValueError:
            return None

    def _on_project_end_changed(self, *_):
        raw = self.project_end_var.get().strip()
        if not raw:
            self._project_end_entry.configure(style="TEntry")
            return
        try:
            from datetime import datetime
            d = datetime.strptime(raw, "%d-%m-%Y").date()
            self._project_end_entry.configure(style="ValidEnd.TEntry")
            maandag = d - timedelta(days=d.weekday())
            iso = maandag.isocalendar()
            self._export_end_iso_year = iso[0]
            self._export_end_week = iso[1]
            if hasattr(self, '_export_end_label'):
                self._update_export_end_label()
            self._propagate_project_dates()
        except ValueError:
            self._project_end_entry.configure(style="InvalidEnd.TEntry")
        self._on_change()

    def _propagate_project_dates(self):
        """Schrijft de huidige project start/einde naar alle opgeslagen weken van dit project."""
        if self._loading:
            return
        project_nr = self.project_nr_var.get().strip() or "P000000"
        project_start = self.project_start_var.get().strip()
        project_end = self.project_end_var.get().strip()
        alle_weken = WeekRapportData.load_all_for_project(project_nr)
        for w in alle_weken:
            changed = False
            if project_start and w.get("project_start") != project_start:
                w["project_start"] = project_start
                changed = True
            if project_end and w.get("project_end") != project_end:
                w["project_end"] = project_end
                changed = True
            if changed:
                WeekRapportData.save(w)

    def _nav_week(self, delta: int):
        self._save_current_week()
        maandag, _ = WeekRapportData.week_dates_from_iso(
            self.current_iso_year, self.current_kalender_week
        )
        nieuwe_maandag = maandag + timedelta(weeks=delta)
        if delta < 0:
            start_monday = self._get_project_start_monday()
            if start_monday and nieuwe_maandag < start_monday:
                return
        if delta > 0:
            end_monday = self._get_project_end_monday()
            if end_monday and nieuwe_maandag > end_monday:
                return
        iso = nieuwe_maandag.isocalendar()
        self.current_iso_year = iso[0]
        self.current_kalender_week = iso[1]
        self._load_week()

    def _update_header_labels(self, data: dict):
        project_week_nr = data.get("project_week_nr", "1")
        kalender_week = data.get("kalender_week_nr", "?")
        week_start = data.get("week_start", "")
        week_end = data.get("week_end", "")

        def fmt(d: str) -> str:
            try:
                from datetime import date
                dt = date.fromisoformat(d)
                return dt.strftime("%d/%m/%Y")
            except Exception:
                return d

        # Weeknr invullen in het invoerveld (alleen als het leeg of anders is)
        if self.project_week_nr_var.get() != str(project_week_nr):
            self.project_week_nr_var.set(str(project_week_nr))

        self._week_label.config(text=f"Kalender week {kalender_week}")
        self._datum_label.config(text=f"{fmt(week_start)} – {fmt(week_end)}")

    # ──────────────────────────────────────────── Laden / Opslaan

    def _load_week(self):
        self._loading = True
        project_nr = self.project_nr_var.get().strip() or "P000000"
        iso_year = self.current_iso_year
        week = self.current_kalender_week

        data = WeekRapportData.load(project_nr, iso_year, week)
        if data is None:
            carried = WeekRapportData.get_latest_werknemers(project_nr)
            data = WeekRapportData.empty_week(
                project_nr, iso_year, week,
                locatie="",
                carried_werknemers=carried
            )

        self._populate_ui(data)
        self._loading = False

        self._update_feestdagen_display(iso_year, week)

        # WEER API TIJDELIJK LOSGEKOPPELD — verwijder onderstaande return om te herkoppelen
        self._weer_status.config(text="Weerdata API losgekoppeld.")
        self._wb_status.config(text="Werkbaarheid API losgekoppeld.")
        return

        # Start achtergrondthread voor weer + werkbaarheid
        week_start_date, week_end_date = WeekRapportData.week_dates_from_iso(iso_year, week)
        locatie = data.get("locatie", "") or "Amsterdam"
        self._weer_status.config(text="Weerdata ophalen...")
        self._wb_status.config(text="Werkbaarheid berekenen...")

        def bg_task():
            weather = None
            result = None
            try:
                from fetch_weather import get_weather_for_period
                weather, _ = get_weather_for_period(locatie, week_start_date, week_end_date)
            except Exception as e:
                self.after(0, self._weer_status.config, {"text": f"Weerdata niet beschikbaar: {e}"})

            try:
                from calculate_workdays import calculate
                result = calculate(week_start_date, week_end_date, locatie, norm="standaard")
            except Exception as e:
                self.after(0, self._wb_status.config, {"text": f"Werkbaarheid niet beschikbaar: {e}"})

            self.after(0, self._populate_weer_werkbaarheid, weather, result)

        threading.Thread(target=bg_task, daemon=True).start()

    def _populate_ui(self, data: dict):
        self._update_header_labels(data)

        # Project start- en einddatum terugzetten vanuit opgeslagen data
        saved_start = data.get("project_start", "")
        if saved_start:
            self.project_start_var.set(saved_start)

        saved_end = data.get("project_end", "")
        if saved_end:
            self.project_end_var.set(saved_end)

        # Weer
        weer = data.get("weer", {})
        for dag in DAYS:
            w = weer.get(dag, {})
            self._weer_vars[dag]["beschrijving"].set(w.get("beschrijving") or "")
            self._weer_vars[dag]["temp_c"].set(
                str(w["temp_c"]) if w.get("temp_c") is not None else ""
            )
            self._weer_vars[dag]["regen_mm"].set(
                str(w["regen_mm"]) if w.get("regen_mm") is not None else ""
            )
            self._weer_vars[dag]["wind_bft"].set(
                str(w["wind_bft"]) if w.get("wind_bft") is not None else ""
            )

        # Handmatig onwerkbaar flags laden vanuit weer data
        for dag in DAYS:
            w = weer.get(dag, {})
            onw = w.get("handmatig_onwerkbaar", False)
            reden = w.get("handmatig_reden", "")
            if dag in self._handmatig_onwerkbaar_vars:
                self._handmatig_onwerkbaar_vars[dag].set(onw)
                self._handmatig_reden_vars[dag].set(reden)
                self._update_onwerkbaar_btn(dag)

        # Werkbaarheid herberekenen op basis van feestdagen + handmatig overrides
        self._recalc_werkbaarheid()

        # Werknemers
        for rij in list(self._werknemer_rijen):
            rij.frame.destroy()
        self._werknemer_rijen.clear()
        for wn in data.get("werknemers", []):
            self._add_werknemer_rij(
                naam=wn.get("naam", ""),
                bedrijf=wn.get("bedrijf", ""),
                functie=wn.get("functie", ""),
                uren=wn.get("uren", {}),
                opmerking=wn.get("opmerking", ""),
            )
        # Altijd minimaal 3 rijen tonen
        while len(self._werknemer_rijen) < 3:
            self._add_werknemer_rij()

        # Werkzaamheden
        wzh = data.get("werkzaamheden", {})
        for dag in DAYS:
            if dag in self._dag_panels:
                self._dag_panels[dag].set_activiteiten(wzh.get(dag, ["", "", ""]))

    _NL_DAGEN = ["Maandag", "Dinsdag", "Woensdag", "Donderdag", "Vrijdag", "Zaterdag", "Zondag"]
    _NL_MAANDEN = ["Januari", "Februari", "Maart", "April", "Mei", "Juni",
                   "Juli", "Augustus", "September", "Oktober", "November", "December"]

    def _update_feestdagen_display(self, iso_year: int, kalender_week: int):
        from dutch_holidays import get_holidays_for_range
        week_start, week_end = WeekRapportData.week_dates_from_iso(iso_year, kalender_week)
        feestdagen = get_holidays_for_range(week_start, week_end)
        if not feestdagen:
            self._feestdag_title.grid_remove()
            self._feestdagen_label.grid_remove()
            return
        regels = []
        for d, naam in sorted(feestdagen.items()):
            dag_naam = self._NL_DAGEN[d.weekday()]
            maand_naam = self._NL_MAANDEN[d.month - 1]
            regels.append(f"{dag_naam} {d.day} {maand_naam} — {naam}")
        self._feestdagen_label.config(text="  |  ".join(regels))
        self._feestdag_title.grid()
        self._feestdagen_label.grid()

    def _toggle_handmatig_onwerkbaar(self, dag: str):
        nieuwe_staat = not self._handmatig_onwerkbaar_vars[dag].get()
        self._handmatig_onwerkbaar_vars[dag].set(nieuwe_staat)
        self._update_onwerkbaar_btn(dag)
        self._recalc_werkbaarheid()
        self._on_change()

    def _update_onwerkbaar_btn(self, dag: str):
        btn = self._handmatig_onwerkbaar_btns.get(dag)
        if not btn or not btn.winfo_exists():
            return
        if self._handmatig_onwerkbaar_vars[dag].get():
            btn.config(text="✕ Onwerkbaar", bg=DANGER, fg="white", relief="flat")
        else:
            btn.config(text="Onwerkbaar", bg="#e9ecef", fg="#495057", relief="groove")

    def _recalc_werkbaarheid(self):
        from dutch_holidays import get_holidays_for_range
        week_start, week_end = WeekRapportData.week_dates_from_iso(
            self.current_iso_year, self.current_kalender_week
        )
        feestdagen = get_holidays_for_range(week_start, week_end)
        feestd_count = sum(1 for d in feestdagen if d.weekday() < 5)
        handmatig_count = sum(
            1 for dag in WORKDAYS
            if dag in self._handmatig_onwerkbaar_vars
            and self._handmatig_onwerkbaar_vars[dag].get()
        )
        werkbare = max(0, 5 - feestd_count - handmatig_count)
        self._wb_vars["werkbare_dagen"].set(str(werkbare))
        self._wb_vars["onwerkbare_dagen"].set(str(feestd_count + handmatig_count))
        self._wb_vars["onwerkbaar_feestdagen"].set(str(feestd_count))
        self._wb_vars["onwerkbaar_weer"].set(str(handmatig_count))

    def _populate_weer_werkbaarheid(self, weather, result):
        from calculate_workdays import STATUS_FEESTDAG, STATUS_ONWERKBAAR_WEER, STATUS_WEEKEND

        # ── Weerdata invullen ──────────────────────────────────────────────
        if weather:
            week_start, _ = WeekRapportData.week_dates_from_iso(
                self.current_iso_year, self.current_kalender_week
            )
            dag_map = {i: d for i, d in enumerate(DAYS)}  # 0→ma, 1→di, ...
            for d, day_data in weather.items():
                weekday_idx = (d - week_start).days
                dag = dag_map.get(weekday_idx)
                if dag is None:
                    continue

                temp = day_data.get("temp_gem_c")
                regen = day_data.get("rain_totaal_mm")
                wind_ms = day_data.get("wind_gem_ms")
                reasons = day_data.get("reasons", [])
                unworkable = day_data.get("unworkable", False)
                mogelijk = day_data.get("mogelijk_onwerkbaar", False)

                # Beschrijving: alleen vullen als leeg
                if not self._weer_vars[dag]["beschrijving"].get():
                    if unworkable and reasons:
                        beschrijving = "Onwerkbaar: " + ", ".join(reasons)
                    elif mogelijk and reasons:
                        beschrijving = "Mogelijk onwerkbaar: " + ", ".join(reasons)
                    elif regen is not None and regen >= 1:
                        beschrijving = f"Neerslag {round(regen, 1)} mm"
                    else:
                        beschrijving = "Droog"
                    self._weer_vars[dag]["beschrijving"].set(beschrijving)

                if not self._weer_vars[dag]["temp_c"].get() and temp is not None:
                    self._weer_vars[dag]["temp_c"].set(str(round(temp, 1)))

                if not self._weer_vars[dag]["regen_mm"].get() and regen is not None:
                    self._weer_vars[dag]["regen_mm"].set(str(round(regen, 1)))

                if not self._weer_vars[dag]["wind_bft"].get() and wind_ms is not None:
                    self._weer_vars[dag]["wind_bft"].set(str(ms_to_beaufort(wind_ms)))

            self._weer_status.config(text="Weerdata geladen.")
        else:
            self._weer_status.config(text="Weerdata niet beschikbaar.")

        # ── Werkbaarheid invullen ──────────────────────────────────────────
        if result:
            werkbare = sum(
                1 for d in result.dagen
                if d.status not in (STATUS_FEESTDAG, STATUS_ONWERKBAAR_WEER, STATUS_WEEKEND)
                and d.datum.weekday() < 5
            )
            feestdagen = sum(
                1 for d in result.dagen
                if d.status == STATUS_FEESTDAG and d.datum.weekday() < 5
            )
            onwerkbaar_weer = sum(
                1 for d in result.dagen
                if d.status == STATUS_ONWERKBAAR_WEER and d.datum.weekday() < 5
            )
            self._wb_vars["werkbare_dagen"].set(str(werkbare))
            self._wb_vars["onwerkbare_dagen"].set(str(feestdagen + onwerkbaar_weer))
            self._wb_vars["onwerkbaar_feestdagen"].set(str(feestdagen))
            self._wb_vars["onwerkbaar_weer"].set(str(onwerkbaar_weer))
            self._wb_status.config(text="Werkbaarheid berekend.")
        else:
            self._wb_status.config(text="Werkbaarheid niet beschikbaar.")

        # Auto-opslaan na laden van weerdata
        self._save_current_week()

    def _save_current_week(self):
        if self._loading:
            return
        project_nr = self.project_nr_var.get().strip() or "P000000"
        iso_year = self.current_iso_year
        week = self.current_kalender_week

        maandag, zondag = WeekRapportData.week_dates_from_iso(iso_year, week)

        # project_week_nr uit het vrij-invoerbare veld lezen
        try:
            project_week_nr = int(self.project_week_nr_var.get())
        except (ValueError, TypeError):
            project_week_nr = WeekRapportData.next_project_week_nr(project_nr)

        # Weerdata uit UI lezen
        weer = {}
        for dag in DAYS:
            def _float(v):
                try:
                    return float(v) if v.strip() else None
                except (ValueError, AttributeError):
                    return None
            def _int(v):
                try:
                    return int(v) if v.strip() else None
                except (ValueError, AttributeError):
                    return None
            weer[dag] = {
                "beschrijving": self._weer_vars[dag]["beschrijving"].get(),
                "temp_c": _float(self._weer_vars[dag]["temp_c"].get()),
                "regen_mm": _float(self._weer_vars[dag]["regen_mm"].get()),
                "wind_bft": _int(self._weer_vars[dag]["wind_bft"].get()),
                "handmatig_onwerkbaar": self._handmatig_onwerkbaar_vars[dag].get(),
                "handmatig_reden": self._handmatig_reden_vars[dag].get(),
            }

        werkbaarheid = {
            "werkbare_dagen": self._safe_int(self._wb_vars["werkbare_dagen"].get()),
            "onwerkbaar_feestdagen": self._safe_int(self._wb_vars["onwerkbaar_feestdagen"].get()),
            "onwerkbaar_weer": self._safe_int(self._wb_vars["onwerkbaar_weer"].get()),
        }

        werknemers = [rij.get_data() for rij in self._werknemer_rijen]

        werkzaamheden = {dag: self._dag_panels[dag].get_activiteiten() for dag in DAYS}

        data = {
            "project_nr": project_nr,
            "project_start": self.project_start_var.get().strip(),
            "project_end": self.project_end_var.get().strip(),
            "project_week_nr": project_week_nr,
            "iso_year": iso_year,
            "kalender_week_nr": week,
            "week_start": maandag.isoformat(),
            "week_end": zondag.isoformat(),
            "locatie": "",
            "weer": weer,
            "werkbaarheid": werkbaarheid,
            "werknemers": werknemers,
            "werkzaamheden": werkzaamheden,
        }

        WeekRapportData.save(data)
        self._werknemers_namen = WeekRapportData.get_all_unique_namen()
        self._werknemers_functies = WeekRapportData.get_all_unique_functies()

        # Nieuw project automatisch registreren
        if project_nr not in self._projecten:
            self._projecten.append(project_nr)
            self._projecten.sort()
            WeekRapportData.save_projecten(self._projecten)

        self._status_label.config(text="Opgeslagen.")
        self.after(3000, lambda: self._status_label.config(text=""))

    @staticmethod
    def _safe_int(val) -> int:
        try:
            return int(val)
        except (ValueError, TypeError):
            return 0

    # ──────────────────────────────────────────── Auto-opslaan

    def _on_change(self):
        if self._loading:
            return
        if self._autosave_after:
            self.after_cancel(self._autosave_after)
        self._autosave_after = self.after(2000, self._save_current_week)

    # ──────────────────────────────────────────── Excel export

    def _export_excel(self):
        self._save_current_week()
        project_nr = self.project_nr_var.get().strip() or "P000000"

        # Alle opgeslagen weken voor dit project ophalen
        alle_weken = WeekRapportData.load_all_for_project(project_nr)
        if not alle_weken:
            messagebox.showwarning("Geen data", "Er zijn nog geen weken opgeslagen voor dit project.", parent=self)
            return

        # Bereik bepalen op basis van project start/einde (indien ingevuld)
        start_monday = self._get_project_start_monday()
        end_monday = self._get_project_end_monday()

        def week_key(data):
            return (data.get("iso_year", 0), data.get("kalender_week_nr", 0))

        def monday_key(d):
            iso = d.isocalendar()
            return (iso[0], iso[1])

        weken = alle_weken
        if start_monday:
            sk = monday_key(start_monday)
            weken = [w for w in weken if week_key(w) >= sk]
        if end_monday:
            ek = monday_key(end_monday)
            weken = [w for w in weken if week_key(w) <= ek]

        if not weken:
            messagebox.showwarning(
                "Geen data",
                "Geen opgeslagen weken gevonden binnen de project-periode.",
                parent=self
            )
            return

        # Sorteren chronologisch op jaar + kalenderweek
        weken.sort(key=lambda w: (w.get("iso_year", 0), w.get("kalender_week_nr", 0)))

        start_kw = weken[0].get("kalender_week_nr", "?")
        eind_kw = weken[-1].get("kalender_week_nr", "?")
        jaar = weken[0].get("iso_year", "")
        standaard_naam = f"Weekrapport_{project_nr}_KW{start_kw}-KW{eind_kw}_{jaar}.xlsx"

        filepath = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=".xlsx",
            initialfile=standaard_naam,
            filetypes=[("Excel-bestand", "*.xlsx"), ("Alle bestanden", "*.*")],
            title=f"Exporteer {len(weken)} week(en) naar Excel"
        )
        if not filepath:
            return

        try:
            from weekrapport_export import export_weeks_to_excel
            from pathlib import Path
            export_weeks_to_excel(weken, Path(filepath))
            messagebox.showinfo(
                "Export geslaagd",
                f"{len(weken)} week(en) geëxporteerd naar:\n{filepath}",
                parent=self
            )
        except Exception as e:
            messagebox.showerror("Exportfout", str(e), parent=self)

    # ──────────────────────────────────────────── Project wisselen

    def _open_project_kiezen(self):
        if not self._projecten:
            messagebox.showinfo(
                "Geen projecten",
                "Er zijn nog geen projecten opgeslagen.\nMaak eerst een nieuw project aan via Extra → Nieuw project.",
                parent=self
            )
            return
        dlg = tk.Toplevel(self)
        dlg.title("Project openen")
        dlg.resizable(False, False)
        dlg.grab_set()

        tk.Label(dlg, text="Kies een project:",
                 font=("Segoe UI", 9, "bold"), padx=16, pady=10).pack(anchor="w")

        frame = tk.Frame(dlg, padx=16)
        frame.pack(fill="both", expand=True)
        lb = tk.Listbox(frame, font=("Segoe UI", 10), width=32, height=min(10, len(self._projecten)),
                        selectmode="single", activestyle="none",
                        selectbackground=PRIMARY, selectforeground="white")
        lb.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(frame, orient="vertical", command=lb.yview)
        lb.config(yscrollcommand=sb.set)
        if len(self._projecten) > 10:
            sb.pack(side="left", fill="y")

        for p in self._projecten:
            lb.insert("end", p)

        # Markeer huidig project
        huidig = self.project_nr_var.get()
        if huidig in self._projecten:
            idx = self._projecten.index(huidig)
            lb.selection_set(idx)
            lb.see(idx)

        def openen():
            sel = lb.curselection()
            if not sel:
                return
            naam = lb.get(sel[0])
            dlg.destroy()
            if naam != self.project_nr_var.get():
                self._switch_project(naam)

        def verwijderen():
            sel = lb.curselection()
            if not sel:
                return
            naam = lb.get(sel[0])
            weken = WeekRapportData.load_all_for_project(naam)
            bevestiging = messagebox.askyesno(
                "Project verwijderen",
                f"Weet je zeker dat je project '{naam}' wilt verwijderen?\n\n"
                f"Dit verwijdert {len(weken)} opgeslagen week(en) permanent.",
                icon="warning", parent=dlg
            )
            if not bevestiging:
                return
            WeekRapportData.delete_project(naam)
            self._projecten = [p for p in self._projecten if p != naam]
            WeekRapportData.save_projecten(self._projecten)
            # Listbox bijwerken
            lb.delete(sel[0])
            # Als huidig project verwijderd: wissel naar eerste overgebleven of leeg
            if naam == self.project_nr_var.get():
                dlg.destroy()
                if self._projecten:
                    self._switch_project(self._projecten[0])
                else:
                    self.project_nr_var.set("P000000")
                    self._load_week()

        lb.bind("<Double-ButtonRelease-1>", lambda _: openen())

        btn_row = tk.Frame(dlg)
        btn_row.pack(pady=10, padx=16, fill="x")
        tk.Button(btn_row, text="Openen", command=openen,
                  font=("Segoe UI", 9, "bold"), bg=PRIMARY, fg="white",
                  relief="flat", padx=14, pady=4, cursor="hand2").pack(side="left", padx=(0, 8))
        tk.Button(btn_row, text="Annuleren", command=dlg.destroy,
                  font=("Segoe UI", 9), relief="flat", padx=10, pady=4,
                  cursor="hand2").pack(side="left")
        tk.Button(btn_row, text="Verwijderen", command=verwijderen,
                  font=("Segoe UI", 9), bg=DANGER, fg="white",
                  relief="flat", padx=10, pady=4, cursor="hand2").pack(side="right")

    def _switch_project(self, naam: str):
        self._save_current_week()
        self.project_nr_var.set(naam)
        # Start/einde wissen — worden ingeladen vanuit opgeslagen data
        self._loading = True
        self.project_start_var.set("")
        self.project_end_var.set("")
        self._loading = False
        # Ga naar de meest recente opgeslagen week, anders huidige week
        weken = WeekRapportData.load_all_for_project(naam)
        if weken:
            laatste = max(weken, key=lambda w: (w.get("iso_year", 0), w.get("kalender_week_nr", 0)))
            self.current_iso_year = laatste["iso_year"]
            self.current_kalender_week = laatste["kalender_week_nr"]
        else:
            iso = date.today().isocalendar()
            self.current_iso_year = iso[0]
            self.current_kalender_week = iso[1]
        self._load_week()

    def _nieuw_project(self):
        dlg = tk.Toplevel(self)
        dlg.title("Nieuw project")
        dlg.resizable(False, False)
        dlg.grab_set()

        tk.Label(dlg, text="Projectnaam / projectnummer:",
                 font=("Segoe UI", 9, "bold"), padx=16, pady=10).pack(anchor="w")
        naam_var = tk.StringVar()
        entry = ttk.Entry(dlg, textvariable=naam_var, width=26, font=("Segoe UI", 10))
        entry.pack(padx=16, pady=(0, 10))
        entry.focus_set()

        def aanmaken():
            naam = naam_var.get().strip()
            if not naam:
                return
            if naam not in self._projecten:
                self._projecten.append(naam)
                self._projecten.sort()
                WeekRapportData.save_projecten(self._projecten)
            dlg.destroy()
            self._switch_project(naam)

        entry.bind("<Return>", lambda _: aanmaken())

        btn_row = tk.Frame(dlg)
        btn_row.pack(pady=(0, 12))
        tk.Button(btn_row, text="Aanmaken", command=aanmaken,
                  font=("Segoe UI", 9, "bold"), bg=PRIMARY, fg="white",
                  relief="flat", padx=14, pady=4, cursor="hand2").pack(side="left", padx=(0, 8))
        tk.Button(btn_row, text="Annuleren", command=dlg.destroy,
                  font=("Segoe UI", 9), relief="flat", padx=10, pady=4,
                  cursor="hand2").pack(side="left")

    # ──────────────────────────────────────────── Venster sluiten

    def _on_close(self):
        self._save_current_week()
        self.destroy()


def main():
    app = WeekRapportApp()
    app.after(500, lambda: _vraag_snelkoppeling(app))
    app.mainloop()


def _vraag_snelkoppeling(root):
    """Toont eenmalig een dialoog om een bureaublad-snelkoppeling aan te maken."""
    import subprocess
    if not getattr(sys, 'frozen', False):
        return
    markeer = os.path.join(os.path.dirname(sys.executable), ".desktop_asked")
    if os.path.exists(markeer):
        return
    try:
        open(markeer, 'w').close()
    except Exception:
        pass
    antwoord = messagebox.askyesno(
        "Snelkoppeling aanmaken",
        "Wil je een snelkoppeling op je bureaublad aanmaken\n"
        "zodat je dit programma makkelijk kunt terugvinden?",
        parent=root
    )
    if antwoord:
        try:
            exe_pad = sys.executable.replace("'", "''")
            r = subprocess.run(
                ["powershell", "-NoProfile", "-Command",
                 "[Environment]::GetFolderPath('Desktop')"],
                capture_output=True, text=True
            )
            bureaublad = r.stdout.strip() or os.path.join(os.path.expanduser("~"), "Desktop")
            snelkoppeling = os.path.join(bureaublad, "Werkzaamheden Weekrapport.lnk").replace("'", "''")
            ps = f"""
$ws = New-Object -ComObject WScript.Shell
$s = $ws.CreateShortcut('{snelkoppeling}')
$s.TargetPath = '{exe_pad}'
$s.IconLocation = '{exe_pad}, 0'
$s.Description = 'Werkzaamheden Weekrapport'
$s.WorkingDirectory = Split-Path '{exe_pad}'
$s.Save()
"""
            result = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                capture_output=True
            )
            if result.returncode == 0:
                messagebox.showinfo(
                    "Snelkoppeling aangemaakt",
                    "De snelkoppeling is aangemaakt op je bureaublad.",
                    parent=root
                )
        except Exception:
            pass


if __name__ == "__main__":
    main()
