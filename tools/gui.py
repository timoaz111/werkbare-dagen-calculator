"""
Werkbare Dagen Calculator — GUI applicatie
Gebouwd met Tkinter voor gebruik op Windows.
"""

import sys
import os
import subprocess
import threading
from datetime import date, datetime

import tkinter as tk
from tkinter import ttk, messagebox

# Zorg dat tools/ in het pad zit
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from calculate_workdays import (
    calculate, STATUS_WERKBAAR, STATUS_WEEKEND,
    STATUS_FEESTDAG, STATUS_MOGELIJK_ONWERKBAAR, STATUS_ONWERKBAAR_WEER
)


from utils import ms_to_beaufort


# Kleuren per status
STATUS_COLORS = {
    STATUS_WERKBAAR: "#d4edda",            # groen
    STATUS_WEEKEND: "#e2e3e5",             # grijs
    STATUS_FEESTDAG: "#fff3cd",            # geel
    STATUS_MOGELIJK_ONWERKBAAR: "#ffe8b0", # oranje
    STATUS_ONWERKBAAR_WEER: "#f8d7da"      # rood
}

STATUS_ICONS = {
    STATUS_WERKBAAR: "✓",
    STATUS_WEEKEND: "—",
    STATUS_FEESTDAG: "🎉",
    STATUS_MOGELIJK_ONWERKBAAR: "⚠",
    STATUS_ONWERKBAAR_WEER: "⛈"
}

APP_TITLE = "Werkbare Dagen Calculator — Bouw"
APP_BG = "#f8f9fa"
PRIMARY = "#0d6efd"
HEADER_BG = "#343a40"
HEADER_FG = "#ffffff"


class WerkbareDagenApp(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title(APP_TITLE)
        self.geometry("1050x750")
        self.minsize(800, 600)
        self.configure(bg=APP_BG)
        self.resizable(True, True)

        # Stel bouwhelm icoon in voor het venster
        try:
            if getattr(sys, 'frozen', False):
                icon_pad = os.path.join(sys._MEIPASS, "tools", "icon.ico")
            else:
                icon_pad = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "icon.ico")
            self.iconbitmap(icon_pad)
        except Exception:
            pass

        self._build_ui()

    # ------------------------------------------------------------------ UI

    def _build_ui(self):
        # Header
        header = tk.Frame(self, bg=HEADER_BG, pady=12)
        header.pack(fill="x")
        tk.Label(
            header, text="Werkbare Dagen Calculator",
            font=("Segoe UI", 18, "bold"),
            bg=HEADER_BG, fg=HEADER_FG
        ).pack()
        tk.Label(
            header, text="CAO Bouwnijverheid — Bft 7 wind / 2mm neerslag / -3°C vorst",
            font=("Segoe UI", 9),
            bg=HEADER_BG, fg="#adb5bd"
        ).pack()

        # Invoerpaneel
        input_frame = tk.LabelFrame(
            self, text="Invoer", font=("Segoe UI", 10, "bold"),
            bg=APP_BG, padx=16, pady=12
        )
        input_frame.pack(fill="x", padx=16, pady=(12, 0))

        # Rij 1: datums + locatie
        row1 = tk.Frame(input_frame, bg=APP_BG)
        row1.pack(fill="x")

        self._make_label(row1, "Startdatum (dd-mm-jjjj):").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.entry_start = self._make_entry(row1, width=14)
        self.entry_start.grid(row=0, column=1, padx=(0, 24))
        self.entry_start.insert(0, "01-01-2024")

        self._make_label(row1, "Einddatum (dd-mm-jjjj):").grid(row=0, column=2, sticky="w", padx=(0, 8))
        self.entry_end = self._make_entry(row1, width=14)
        self.entry_end.grid(row=0, column=3, padx=(0, 24))
        self.entry_end.insert(0, date.today().strftime("%d-%m-%Y"))

        self._make_label(row1, "Locatie (stad):").grid(row=0, column=4, sticky="w", padx=(0, 8))
        self.entry_location = self._make_entry(row1, width=20)
        self.entry_location.grid(row=0, column=5, padx=(0, 24))
        self.entry_location.insert(0, "Amsterdam")

        self.btn_calculate = tk.Button(
            row1, text="Bereken", command=self._start_calculation,
            font=("Segoe UI", 10, "bold"),
            bg=PRIMARY, fg="white", relief="flat",
            padx=18, pady=6, cursor="hand2",
            activebackground="#0b5ed7", activeforeground="white"
        )
        self.btn_calculate.grid(row=0, column=6, padx=(8, 0))

        # Rij 2: werksoort
        row2 = tk.Frame(input_frame, bg=APP_BG)
        row2.pack(fill="x", pady=(10, 0))

        self._make_label(row2, "Werksoort:").pack(side="left", padx=(0, 12))

        self.norm_var = tk.StringVar(value="standaard")
        tk.Radiobutton(
            row2, text="Standaard bouw  (wind \u2265 Bft 7 / \u226513,9 m/s)",
            variable=self.norm_var, value="standaard",
            font=("Segoe UI", 9), bg=APP_BG, activebackground=APP_BG
        ).pack(side="left", padx=(0, 20))
        tk.Radiobutton(
            row2, text="Hijswerkzaamheden  (wind \u2265 Bft 6 / \u226510,8 m/s)",
            variable=self.norm_var, value="hijswerk",
            font=("Segoe UI", 9), bg=APP_BG, activebackground=APP_BG
        ).pack(side="left")

        # Voortgangsbalk
        self.progress_var = tk.StringVar(value="")
        self.lbl_progress = tk.Label(
            input_frame, textvariable=self.progress_var,
            font=("Segoe UI", 9, "italic"), bg=APP_BG, fg="#6c757d"
        )
        self.lbl_progress.pack(anchor="w", pady=(8, 0))

        # Samenvattingspaneel
        summary_outer = tk.Frame(self, bg=APP_BG)
        summary_outer.pack(fill="x", padx=16, pady=(10, 0))

        summary_title_row = tk.Frame(summary_outer, bg=APP_BG)
        summary_title_row.pack(fill="x")
        tk.Label(
            summary_title_row, text="Samenvatting",
            font=("Segoe UI", 10, "bold"), bg=APP_BG
        ).pack(side="left")
        tk.Button(
            summary_title_row, text="ⓘ",
            font=("Segoe UI", 10), bg=APP_BG, fg="#0d6efd",
            relief="flat", cursor="hand2", bd=0,
            command=self._show_definition
        ).pack(side="left", padx=(6, 0))

        summary_frame = tk.LabelFrame(
            summary_outer, text="", font=("Segoe UI", 10, "bold"),
            bg=APP_BG, padx=16, pady=10
        )
        summary_frame.pack(fill="x")

        self.summary_vars = {}
        # (key, label, rij, kolom)
        summary_layout = [
            ("locatie",                "Locatie:",                       0, 0),
            ("periode",                "Periode:",                       0, 2),
            ("totaal",                 "Totaal kalenderdagen:",          0, 4),
            ("werkbaar",               "Werkbare dagen:",                1, 0),
            ("niet_werkbaar",          "Niet werkbare dagen:",           1, 2),
            ("weekenden",              "  ↳ Weekenden:",                 2, 0),
            ("feestdagen",             "  ↳ Feestdagen:",                2, 2),
            ("onwerkbaar_weer",        "  ↳ Onwerkbaar weer:",           3, 0),
            ("mogelijk_onwerkbaar",    "  ↳ Mogelijk niet-werkbaar:",    3, 2),
            ("niet_werkbaar_ex_weekend","Niet werkbaar (excl. weekend):", 4, 0),
            ("weerverlet",             "Weerverlet dagen:",              4, 2),
        ]
        for key, label, row, col in summary_layout:
            tk.Label(
                summary_frame, text=label,
                font=("Segoe UI", 9, "bold"), bg=APP_BG, anchor="w"
            ).grid(row=row, column=col, sticky="w", padx=(0, 4), pady=2)
            var = tk.StringVar(value="—")
            self.summary_vars[key] = var
            tk.Label(
                summary_frame, textvariable=var,
                font=("Segoe UI", 9), bg=APP_BG, anchor="w", width=20
            ).grid(row=row, column=col + 1, sticky="w", padx=(0, 24), pady=2)

        # Legenda
        legend_frame = tk.Frame(self, bg=APP_BG)
        legend_frame.pack(fill="x", padx=16, pady=(8, 0))

        for status, color in STATUS_COLORS.items():
            icon = STATUS_ICONS[status]
            box = tk.Frame(legend_frame, bg=color, width=14, height=14, relief="solid", bd=1)
            box.pack(side="left", padx=(0, 4))
            tk.Label(
                legend_frame, text=f"{icon} {status}",
                font=("Segoe UI", 9), bg=APP_BG
            ).pack(side="left", padx=(0, 16))

        # Detailtabel
        table_frame = tk.LabelFrame(
            self, text="Details per dag", font=("Segoe UI", 10, "bold"),
            bg=APP_BG, padx=8, pady=8
        )
        table_frame.pack(fill="both", expand=True, padx=16, pady=(8, 16))

        # Zoekbalk
        search_row = tk.Frame(table_frame, bg=APP_BG)
        search_row.pack(fill="x", pady=(0, 6))
        tk.Label(search_row, text="Filter:", font=("Segoe UI", 9), bg=APP_BG).pack(side="left", padx=(0, 6))
        self.filter_var = tk.StringVar()
        self.filter_var.trace("w", self._apply_filter)
        filter_entry = ttk.Entry(search_row, textvariable=self.filter_var, width=20)
        filter_entry.pack(side="left", padx=(0, 12))

        self.filter_status = tk.StringVar(value="Alle")
        status_options = ["Alle", STATUS_WERKBAAR, STATUS_WEEKEND, STATUS_FEESTDAG, STATUS_MOGELIJK_ONWERKBAAR, STATUS_ONWERKBAAR_WEER]
        ttk.Combobox(
            search_row, textvariable=self.filter_status,
            values=status_options, state="readonly", width=22
        ).pack(side="left")
        self.filter_status.trace("w", self._apply_filter)

        # Tabel
        columns = ("datum", "dag", "status", "reden", "wind", "regen_tot", "regen_max", "temp_gem")
        self.tree = ttk.Treeview(
            table_frame, columns=columns, show="headings",
            selectmode="browse"
        )
        self.tree.heading("datum", text="Datum ↕", command=self._sort_datum)
        self.tree.heading("dag", text="Dag")
        self.tree.heading("status", text="Status")
        self.tree.heading("reden", text="Reden / Opmerking")
        self.tree.heading("wind", text="Gem. wind (Bft)")
        self.tree.heading("regen_tot", text="Totale neerslag (mm)")
        self.tree.heading("regen_max", text="Max neerslag (mm/u)")
        self.tree.heading("temp_gem", text="Gem. temp. (°C)")

        self.tree.column("datum", width=100, anchor="center")
        self.tree.column("dag", width=90, anchor="center")
        self.tree.column("status", width=155, anchor="center")
        self.tree.column("reden", width=230, anchor="w")
        self.tree.column("wind", width=105, anchor="center")
        self.tree.column("regen_tot", width=135, anchor="center")
        self.tree.column("regen_max", width=135, anchor="center")
        self.tree.column("temp_gem", width=110, anchor="center")

        # Kleur-tags
        for status, color in STATUS_COLORS.items():
            self.tree.tag_configure(status, background=color)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        # Sla alle dag-data op voor filtering
        self._all_days = []
        self._datum_oplopend = True  # sorteervolgorde datum

    # ------------------------------------------------------------------ Helpers

    def _show_definition(self):
        popup = tk.Toplevel(self)
        popup.title("Definitie Niet-Werkbare Dag")
        popup.geometry("480x420")
        popup.resizable(True, True)
        popup.configure(bg=APP_BG)
        popup.grab_set()

        tk.Label(
            popup, text="Definitie Niet-Werkbare Dag",
            font=("Segoe UI", 13, "bold"), bg=APP_BG
        ).pack(pady=(20, 4))
        tk.Label(
            popup, text="CAO Bouwnijverheid",
            font=("Segoe UI", 9, "italic"), bg=APP_BG, fg="#6c757d"
        ).pack()

        ttk.Separator(popup, orient="horizontal").pack(fill="x", padx=20, pady=12)

        # Scrollbaar tekstgebied
        text_frame = tk.Frame(popup, bg=APP_BG)
        text_frame.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        scrollbar = ttk.Scrollbar(text_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        tekstvak = tk.Text(
            text_frame, font=("Segoe UI", 10), bg=APP_BG,
            relief="flat", wrap="word", cursor="arrow",
            yscrollcommand=scrollbar.set,
            padx=4, pady=4, state="normal"
        )
        tekstvak.pack(fill="both", expand=True)
        scrollbar.config(command=tekstvak.yview)

        norm = self.norm_var.get()
        if norm == "hijswerk":
            wind_regel = "  \U0001f32c  Wind       \u2265  Beaufort 6  (\u2265 10,8 m/s)  \u2014 hijsnorm"
        else:
            wind_regel = "  \U0001f32c  Wind       \u2265  Beaufort 7  (\u2265 13,9 m/s)"

        tekst = (
            "Bron: Bouw & Infra CAO \u2014 Richtlijnen buitenwerk\n\n"
            "NEERSLAG\n"
            "Een dag is niet-werkbaar door regen als:\n"
            "  \U0001f327  Totale neerslag > 20 mm tijdens de werkdag,\n"
            "       EN/OF\n"
            "  \U0001f327  Neerslag > 5 mm/u gedurende 2 uur of meer.\n\n"
            "WIND\n"
            f"  \U0001f32c  {wind_regel[4:]}\n"
            "       Niet-werkbaar bij \u2265 5 uur binnen de werkdag.\n\n"
            "VORST\n"
            "  \U0001f321  Temperatuur \u2264 -3 \u00b0C\n"
            "       Niet-werkbaar bij \u2265 5 uur binnen de werkdag.\n\n"
            "VASTE NIET-WERKBARE DAGEN\n"
            "  \U0001f4c5  Zaterdag en zondag (weekend)\n"
            "  \U0001f389  Offici\u00eble Nederlandse feestdagen\n"
            "       (Nieuwjaar, Pasen, Koningsdag, Bevrijdingsdag,\n"
            "        Hemelvaartsdag, Pinksteren, Kerst)\n\n"
            "Werkdag = 07:00 \u2013 17:00 uur (10 uur)"
        )
        tekstvak.insert("1.0", tekst)
        tekstvak.config(state="disabled")

        tk.Button(
            popup, text="Sluiten", command=popup.destroy,
            font=("Segoe UI", 10), bg=PRIMARY, fg="white",
            relief="flat", padx=20, pady=6, cursor="hand2"
        ).pack(pady=(0, 20))

    def _make_label(self, parent, text):
        return tk.Label(parent, text=text, font=("Segoe UI", 9), bg=APP_BG)

    def _make_entry(self, parent, width=16):
        e = ttk.Entry(parent, font=("Segoe UI", 10), width=width)
        return e

    def _set_progress(self, msg: str):
        self.progress_var.set(msg)
        self.update_idletasks()

    # ------------------------------------------------------------------ Berekening

    def _start_calculation(self):
        start_str = self.entry_start.get().strip()
        end_str = self.entry_end.get().strip()
        location = self.entry_location.get().strip()

        # Validatie
        try:
            start = datetime.strptime(start_str, "%d-%m-%Y").date()
        except ValueError:
            messagebox.showerror("Fout", "Startdatum is ongeldig. Gebruik dd-mm-jjjj.")
            return

        try:
            end = datetime.strptime(end_str, "%d-%m-%Y").date()
        except ValueError:
            messagebox.showerror("Fout", "Einddatum is ongeldig. Gebruik dd-mm-jjjj.")
            return

        if end < start:
            messagebox.showerror("Fout", "Einddatum moet na de startdatum liggen.")
            return

        if not location:
            messagebox.showerror("Fout", "Voer een locatie in.")
            return

        self.btn_calculate.config(state="disabled", text="Bezig...")
        self._clear_table()
        self._all_days = []

        def run():
            try:
                result = calculate(
                    start, end, location,
                    progress_callback=lambda msg: self.after(0, self._set_progress, msg),
                    norm=self.norm_var.get()
                )
                self.after(0, self._show_result, result)
            except ValueError as e:
                self.after(0, messagebox.showerror, "Fout", str(e))
                self.after(0, self._reset_button)
            except ConnectionError as e:
                self.after(0, messagebox.showerror, "Verbindingsfout", str(e))
                self.after(0, self._reset_button)
            except Exception as e:
                self.after(0, messagebox.showerror, "Onverwachte fout", str(e))
                self.after(0, self._reset_button)

        threading.Thread(target=run, daemon=True).start()

    def _reset_button(self):
        self.btn_calculate.config(state="normal", text="Bereken")
        self._set_progress("")

    def _clear_table(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

    def _show_result(self, result):
        # Samenvatting
        self.summary_vars["locatie"].set(result.locatie)
        self.summary_vars["periode"].set(
            f"{result.start.strftime('%d-%m-%Y')} t/m {result.eind.strftime('%d-%m-%Y')}"
        )
        self.summary_vars["totaal"].set(str(result.totaal))
        self.summary_vars["werkbaar"].set(str(result.werkbaar))
        self.summary_vars["niet_werkbaar"].set(str(result.niet_werkbaar))
        self.summary_vars["weekenden"].set(str(result.weekenden))
        self.summary_vars["feestdagen"].set(str(result.feestdagen))
        self.summary_vars["onwerkbaar_weer"].set(str(result.onwerkbaar_weer))
        self.summary_vars["niet_werkbaar_ex_weekend"].set(
            str(result.feestdagen + result.onwerkbaar_weer)
        )
        self.summary_vars["mogelijk_onwerkbaar"].set(str(result.mogelijk_onwerkbaar))
        self.summary_vars["weerverlet"].set(str(result.onwerkbaar_weer))

        # Sla alle dagen op voor filtering
        def fmt_wind(v):  return f"Bft {ms_to_beaufort(v)}" if v is not None else "—"
        def fmt_mm(v):    return f"{v} mm" if v is not None else "—"
        def fmt_temp(v):  return f"{v} °C" if v is not None else "—"

        self._all_days = [
            (
                d.datum.strftime("%d-%m-%Y"),
                d.dag_naam,
                d.status,
                d.reden,
                fmt_wind(d.wind_gem_ms),
                fmt_mm(d.rain_totaal_mm),
                fmt_mm(d.rain_max_mm),
                fmt_temp(d.temp_gem_c),
            )
            for d in result.dagen
        ]

        self._apply_filter()
        self._set_progress(f"Klaar — {result.werkbaar} werkbare dagen gevonden.")
        self._reset_button()

    def _sort_datum(self):
        self._datum_oplopend = not self._datum_oplopend
        pijl = "↑" if self._datum_oplopend else "↓"
        self.tree.heading("datum", text=f"Datum {pijl}")
        self._apply_filter()

    def _apply_filter(self, *_):
        text_filter = self.filter_var.get().strip().lower()
        status_filter = self.filter_status.get()

        dagen = sorted(
            self._all_days,
            key=lambda r: datetime.strptime(r[0], "%d-%m-%Y"),
            reverse=not self._datum_oplopend
        )

        self._clear_table()

        for datum, dag, status, reden, wind, regen_tot, regen_max, temp_gem in dagen:
            if status_filter != "Alle" and status != status_filter:
                continue
            if text_filter and text_filter not in datum.lower() and \
               text_filter not in dag.lower() and \
               text_filter not in reden.lower():
                continue
            icon = STATUS_ICONS.get(status, "")
            self.tree.insert(
                "", "end",
                values=(datum, dag, f"{icon}  {status}", reden, wind, regen_tot, regen_max, temp_gem),
                tags=(status,)
            )


def _maak_snelkoppeling():
    """Maakt een snelkoppeling op het bureaublad aan (alleen als .exe)."""
    if not getattr(sys, 'frozen', False):
        return False
    try:
        exe_pad = sys.executable.replace("'", "''")
        # Haal het echte bureaublad-pad op via PowerShell (werkt ook met OneDrive)
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "[Environment]::GetFolderPath('Desktop')"],
            capture_output=True, text=True
        )
        bureaublad = r.stdout.strip() or os.path.join(os.path.expanduser("~"), "Desktop")
        snelkoppeling = os.path.join(bureaublad, "Werkbare Dagen Calculator.lnk").replace("'", "''")
        ps = f"""
$ws = New-Object -ComObject WScript.Shell
$s = $ws.CreateShortcut('{snelkoppeling}')
$s.TargetPath = '{exe_pad}'
$s.IconLocation = '{exe_pad}, 0'
$s.Description = 'Werkbare Dagen Calculator'
$s.WorkingDirectory = Split-Path '{exe_pad}'
$s.Save()
"""
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
            capture_output=True
        )
        return result.returncode == 0
    except Exception:
        return False


def _vraag_snelkoppeling(root):
    """Toont eenmalig een dialoog om een bureaublad-snelkoppeling aan te maken."""
    if not getattr(sys, 'frozen', False):
        return  # Alleen in gebouwde .exe

    # Markeerbestand — als dit bestaat, al eerder gevraagd
    markeer = os.path.join(os.path.dirname(sys.executable), ".desktop_asked")
    if os.path.exists(markeer):
        return

    # Markeer direct zodat het maar één keer gevraagd wordt
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
        succes = _maak_snelkoppeling()
        if succes:
            messagebox.showinfo(
                "Snelkoppeling aangemaakt",
                "De snelkoppeling is aangemaakt op je bureaublad.",
                parent=root
            )


def main():
    root = tk.Tk()
    root.withdraw()
    app = WerkbareDagenApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
