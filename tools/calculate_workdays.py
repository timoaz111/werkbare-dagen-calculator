"""
Berekent werkbare en niet-werkbare dagen voor een bouwproject.
Combineert feestdagenlogica en weersdata.
"""

from datetime import date, timedelta
from dataclasses import dataclass, field
from typing import Optional

from dutch_holidays import get_holidays_for_range
from fetch_weather import get_weather_for_period


DAY_NAMES_NL = {
    0: "Maandag", 1: "Dinsdag", 2: "Woensdag",
    3: "Donderdag", 4: "Vrijdag", 5: "Zaterdag", 6: "Zondag"
}

STATUS_WERKBAAR = "Werkbaar"
STATUS_WEEKEND = "Weekend"
STATUS_FEESTDAG = "Feestdag"
STATUS_MOGELIJK_ONWERKBAAR = "Mogelijk niet-werkbaar"
STATUS_ONWERKBAAR_WEER = "Onwerkbaar weer"


@dataclass
class DayResult:
    datum: date
    dag_naam: str
    status: str
    reden: str = ""
    wind_gem_ms: Optional[float] = None
    temp_gem_c: Optional[float] = None
    rain_uren_zwaar: Optional[int] = None    # uren > 5 mm/u
    rain_totaal_mm: Optional[float] = None   # totale neerslag werkdag
    rain_max_mm: Optional[float] = None      # hoogste mm/u op werkdag


@dataclass
class PeriodResult:
    locatie: str
    start: date
    eind: date
    dagen: list[DayResult] = field(default_factory=list)

    @property
    def totaal(self) -> int:
        return len(self.dagen)

    @property
    def werkbaar(self) -> int:
        return sum(1 for d in self.dagen if d.status == STATUS_WERKBAAR)

    @property
    def niet_werkbaar(self) -> int:
        return self.totaal - self.werkbaar

    @property
    def weekenden(self) -> int:
        return sum(1 for d in self.dagen if d.status == STATUS_WEEKEND)

    @property
    def feestdagen(self) -> int:
        return sum(1 for d in self.dagen if d.status == STATUS_FEESTDAG)

    @property
    def mogelijk_onwerkbaar(self) -> int:
        return sum(1 for d in self.dagen if d.status == STATUS_MOGELIJK_ONWERKBAAR)

    @property
    def onwerkbaar_weer(self) -> int:
        return sum(1 for d in self.dagen if d.status == STATUS_ONWERKBAAR_WEER)


def calculate(
    start: date,
    end: date,
    location: str,
    progress_callback=None,
    norm: str = "standaard"
) -> PeriodResult:
    """
    Berekent werkbare dagen voor de opgegeven periode en locatie.

    progress_callback: optionele functie(message: str) voor voortgangsrapportage
    """
    def log(msg):
        if progress_callback:
            progress_callback(msg)

    if end < start:
        raise ValueError("Einddatum moet na de startdatum liggen.")

    today = date.today()
    if start > today:
        raise ValueError("Startdatum ligt in de toekomst. Weersdata is niet beschikbaar.")

    # Kap einddatum op vandaag voor weersdata
    weather_end = min(end, today)
    future_start = today + timedelta(days=1) if end > today else None

    log("Feestdagen ophalen...")
    holidays = get_holidays_for_range(start, end)

    log(f"Weersdata ophalen voor {location}...")
    weather_per_day, display_name = get_weather_for_period(location, start, weather_end, norm=norm)

    log("Dagen berekenen...")
    result = PeriodResult(locatie=display_name, start=start, eind=end)

    current = start
    while current <= end:
        weekday = current.weekday()
        dag_naam = DAY_NAMES_NL[weekday]

        if weekday >= 5:
            # Zaterdag of zondag
            result.dagen.append(DayResult(current, dag_naam, STATUS_WEEKEND))

        elif current in holidays:
            # Officiële feestdag (weekdag)
            naam = holidays[current]
            result.dagen.append(DayResult(current, dag_naam, STATUS_FEESTDAG, naam))

        elif future_start and current >= future_start:
            # Toekomstige werkdag — geen weersdata beschikbaar
            result.dagen.append(DayResult(
                current, dag_naam, STATUS_WERKBAAR,
                "Toekomstige datum (geen weersdata)"
            ))

        elif current in weather_per_day and weather_per_day[current]["unworkable"]:
            # Onwerkbaar door weer
            w = weather_per_day[current]
            reasons = ", ".join(w["reasons"])
            result.dagen.append(DayResult(
                current, dag_naam, STATUS_ONWERKBAAR_WEER, reasons,
                wind_gem_ms=w.get("wind_gem_ms"),
                temp_gem_c=w.get("temp_gem_c"),
                rain_uren_zwaar=w.get("rain_hours_zwaar"),
                rain_totaal_mm=w.get("rain_totaal_mm"),
                rain_max_mm=w.get("rain_max_mm"),
            ))

        elif current in weather_per_day and weather_per_day[current].get("mogelijk_onwerkbaar"):
            # Mogelijk niet-werkbaar (10–20 mm totaal)
            w = weather_per_day[current]
            reden = f"Neerslag {w.get('rain_totaal_mm', 0)} mm (10–20 mm grens)"
            result.dagen.append(DayResult(
                current, dag_naam, STATUS_MOGELIJK_ONWERKBAAR, reden,
                wind_gem_ms=w.get("wind_gem_ms"),
                temp_gem_c=w.get("temp_gem_c"),
                rain_uren_zwaar=w.get("rain_hours_zwaar"),
                rain_totaal_mm=w.get("rain_totaal_mm"),
                rain_max_mm=w.get("rain_max_mm"),
            ))

        else:
            # Werkbare dag
            w = weather_per_day.get(current, {})
            result.dagen.append(DayResult(
                current, dag_naam, STATUS_WERKBAAR,
                wind_gem_ms=w.get("wind_gem_ms"),
                temp_gem_c=w.get("temp_gem_c"),
                rain_uren_zwaar=w.get("rain_hours_zwaar"),
                rain_totaal_mm=w.get("rain_totaal_mm"),
                rain_max_mm=w.get("rain_max_mm"),
            ))

        current += timedelta(days=1)

    return result


if __name__ == "__main__":
    from datetime import date

    def print_progress(msg):
        print(f"  > {msg}")

    start = date(2024, 1, 1)
    end = date(2024, 3, 31)

    print(f"Berekening voor Amsterdam, {start} t/m {end}")
    result = calculate(start, end, "Amsterdam", print_progress)

    print(f"\nLocatie: {result.locatie}")
    print(f"Totaal dagen: {result.totaal}")
    print(f"Werkbaar: {result.werkbaar}")
    print(f"Niet werkbaar: {result.niet_werkbaar}")
    print(f"  - Weekenden: {result.weekenden}")
    print(f"  - Feestdagen: {result.feestdagen}")
    print(f"  - Onwerkbaar weer: {result.onwerkbaar_weer}")

    print("\nOnwerkbare werkdagen:")
    for d in result.dagen:
        if d.status in (STATUS_FEESTDAG, STATUS_ONWERKBAAR_WEER):
            print(f"  {d.datum.strftime('%d-%m-%Y')} {d.dag_naam}: {d.status} — {d.reden}")
