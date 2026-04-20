"""
Data model en JSON persistentie voor het Werkzaamheden Weekrapport.
Elke week wordt opgeslagen als een apart JSON-bestand in de data/ map.
"""

import json
from datetime import date, timedelta
from pathlib import Path


DAYS = ["ma", "di", "wo", "do", "vr", "za", "zo"]


class WeekRapportData:
    # Pad relatief aan de tools/ map → projectroot/data/
    DATA_DIR = Path(__file__).parent.parent / "data"

    @staticmethod
    def filepath(project_nr: str, iso_year: int, kalender_week: int) -> Path:
        return WeekRapportData.DATA_DIR / f"weekrapport_{project_nr}_{iso_year}_W{kalender_week:02d}.json"

    @staticmethod
    def load(project_nr: str, iso_year: int, kalender_week: int) -> dict | None:
        fp = WeekRapportData.filepath(project_nr, iso_year, kalender_week)
        if not fp.exists():
            return None
        with open(fp, "r", encoding="utf-8") as f:
            return json.load(f)

    @staticmethod
    def save(data: dict) -> None:
        WeekRapportData.DATA_DIR.mkdir(parents=True, exist_ok=True)
        fp = WeekRapportData.filepath(
            data["project_nr"], data["iso_year"], data["kalender_week_nr"]
        )
        with open(fp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    @staticmethod
    def load_all_for_project(project_nr: str) -> list[dict]:
        if not WeekRapportData.DATA_DIR.exists():
            return []
        results = []
        for fp in sorted(WeekRapportData.DATA_DIR.glob(f"weekrapport_{project_nr}_*.json")):
            with open(fp, "r", encoding="utf-8") as f:
                results.append(json.load(f))
        return results

    @staticmethod
    def get_latest_werknemers(project_nr: str) -> list[dict]:
        """Geeft naam/bedrijf/functie van werknemers uit de meest recente week."""
        weken = WeekRapportData.load_all_for_project(project_nr)
        if not weken:
            return []
        # Sorteer op iso_year + kalender_week_nr
        weken.sort(key=lambda w: (w.get("iso_year", 0), w.get("kalender_week_nr", 0)))
        laatste = weken[-1]
        werknemers = laatste.get("werknemers", [])
        return [
            {"naam": w.get("naam", ""), "bedrijf": w.get("bedrijf", ""), "functie": w.get("functie", "")}
            for w in werknemers
        ]

    @staticmethod
    def next_project_week_nr(project_nr: str) -> int:
        weken = WeekRapportData.load_all_for_project(project_nr)
        if not weken:
            return 1
        return max(w.get("project_week_nr", 0) for w in weken) + 1

    BEDRIJVEN_FILE    = Path(__file__).parent.parent / "data" / "bedrijven.json"
    INSTELLINGEN_FILE = Path(__file__).parent.parent / "data" / "instellingen.json"
    PROJECTEN_FILE    = Path(__file__).parent.parent / "data" / "projecten.json"

    @staticmethod
    def load_projecten() -> list[str]:
        if not WeekRapportData.PROJECTEN_FILE.exists():
            return []
        with open(WeekRapportData.PROJECTEN_FILE, "r", encoding="utf-8") as f:
            return json.load(f)

    @staticmethod
    def save_projecten(projecten: list[str]) -> None:
        WeekRapportData.DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(WeekRapportData.PROJECTEN_FILE, "w", encoding="utf-8") as f:
            json.dump(sorted(set(p.strip() for p in projecten if p.strip())),
                      f, ensure_ascii=False, indent=2)

    @staticmethod
    def delete_project(project_nr: str) -> int:
        """Verwijdert alle weekbestanden voor dit project. Geeft aantal verwijderde bestanden terug."""
        if not WeekRapportData.DATA_DIR.exists():
            return 0
        bestanden = list(WeekRapportData.DATA_DIR.glob(f"weekrapport_{project_nr}_*.json"))
        for fp in bestanden:
            fp.unlink()
        return len(bestanden)

    @staticmethod
    def load_instellingen() -> dict:
        if not WeekRapportData.INSTELLINGEN_FILE.exists():
            return {"weer_za_zo_verbergen": False, "werknemers_za_zo_verbergen": False}
        with open(WeekRapportData.INSTELLINGEN_FILE, "r", encoding="utf-8") as f:
            return json.load(f)

    @staticmethod
    def save_instellingen(instellingen: dict) -> None:
        WeekRapportData.DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(WeekRapportData.INSTELLINGEN_FILE, "w", encoding="utf-8") as f:
            json.dump(instellingen, f, ensure_ascii=False, indent=2)

    @staticmethod
    def get_all_unique_namen() -> list[str]:
        if not WeekRapportData.DATA_DIR.exists():
            return []
        namen = set()
        for fp in WeekRapportData.DATA_DIR.glob("weekrapport_*.json"):
            with open(fp, "r", encoding="utf-8") as f:
                data = json.load(f)
            for w in data.get("werknemers", []):
                naam = w.get("naam", "").strip()
                if naam:
                    namen.add(naam)
        return sorted(namen)

    @staticmethod
    def get_all_unique_functies() -> list[str]:
        if not WeekRapportData.DATA_DIR.exists():
            return []
        functies = set()
        for fp in WeekRapportData.DATA_DIR.glob("weekrapport_*.json"):
            with open(fp, "r", encoding="utf-8") as f:
                data = json.load(f)
            for w in data.get("werknemers", []):
                functie = w.get("functie", "").strip()
                if functie:
                    functies.add(functie)
        return sorted(functies)

    @staticmethod
    def load_bedrijven() -> list[str]:
        if not WeekRapportData.BEDRIJVEN_FILE.exists():
            return []
        with open(WeekRapportData.BEDRIJVEN_FILE, "r", encoding="utf-8") as f:
            return json.load(f)

    @staticmethod
    def save_bedrijven(bedrijven: list[str]) -> None:
        WeekRapportData.DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(WeekRapportData.BEDRIJVEN_FILE, "w", encoding="utf-8") as f:
            json.dump(sorted(set(b.strip() for b in bedrijven if b.strip())), f, ensure_ascii=False, indent=2)

    @staticmethod
    def week_dates_from_iso(iso_year: int, kalender_week: int) -> tuple[date, date]:
        """Geeft (maandag, zondag) voor de opgegeven ISO-week."""
        maandag = date.fromisocalendar(iso_year, kalender_week, 1)
        zondag = maandag + timedelta(days=6)
        return maandag, zondag

    @staticmethod
    def empty_week(project_nr: str, iso_year: int, kalender_week: int,
                   locatie: str = "", carried_werknemers: list[dict] | None = None) -> dict:
        """Maakt een leeg week-data dict aan, optioneel met doorgevoerde werknemers."""
        maandag, zondag = WeekRapportData.week_dates_from_iso(iso_year, kalender_week)
        project_week_nr = WeekRapportData.next_project_week_nr(project_nr)

        werknemers = []
        for w in (carried_werknemers or []):
            werknemers.append({
                "naam": w.get("naam", ""),
                "bedrijf": w.get("bedrijf", ""),
                "functie": w.get("functie", ""),
                "uren": {dag: 0 for dag in DAYS}
            })

        return {
            "project_nr": project_nr,
            "project_week_nr": project_week_nr,
            "iso_year": iso_year,
            "kalender_week_nr": kalender_week,
            "week_start": maandag.isoformat(),
            "week_end": zondag.isoformat(),
            "locatie": locatie,
            "weer": {
                dag: {"beschrijving": "", "temp_c": None, "regen_mm": None, "wind_bft": None}
                for dag in DAYS
            },
            "werkbaarheid": {
                "werkbare_dagen": 0,
                "onwerkbaar_feestdagen": 0,
                "onwerkbaar_weer": 0
            },
            "werknemers": werknemers,
            "werkzaamheden": {dag: ["", "", ""] for dag in DAYS}
        }
