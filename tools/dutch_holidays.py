"""
Berekent officiële Nederlandse feestdagen voor een gegeven jaar.
Gebaseerd op de CAO Bouwnijverheid definitie van niet-werkbare feestdagen.
"""

from datetime import date, timedelta


def easter_sunday(year: int) -> date:
    """Berekent Eerste Paasdag via het algoritme van Butcher/Anonymous Gregorian."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def get_dutch_holidays(year: int) -> dict[date, str]:
    """
    Geeft een dictionary van alle officiële Nederlandse feestdagen voor een jaar.
    Keys zijn date objecten, values zijn de naam van de feestdag.
    """
    holidays = {}

    # Vaste feestdagen
    holidays[date(year, 1, 1)] = "Nieuwjaarsdag"
    holidays[date(year, 12, 25)] = "Eerste Kerstdag"
    holidays[date(year, 12, 26)] = "Tweede Kerstdag"

    # Koningsdag: 27 april, maar als dat een zondag is → 26 april
    koningsdag = date(year, 4, 27)
    if koningsdag.weekday() == 6:  # zondag
        koningsdag = date(year, 4, 26)
    holidays[koningsdag] = "Koningsdag"

    # Bevrijdingsdag: 5 mei
    holidays[date(year, 5, 5)] = "Bevrijdingsdag"

    # Variabele feestdagen op basis van Pasen
    pasen = easter_sunday(year)
    holidays[pasen - timedelta(days=2)] = "Goede Vrijdag"
    holidays[pasen] = "Eerste Paasdag"
    holidays[pasen + timedelta(days=1)] = "Tweede Paasdag"
    holidays[pasen + timedelta(days=39)] = "Hemelvaartsdag"
    holidays[pasen + timedelta(days=49)] = "Eerste Pinksterdag"
    holidays[pasen + timedelta(days=50)] = "Tweede Pinksterdag"

    return holidays


def get_holidays_for_range(start: date, end: date) -> dict[date, str]:
    """Geeft alle feestdagen voor alle jaren in de opgegeven periode."""
    holidays = {}
    for year in range(start.year, end.year + 1):
        holidays.update(get_dutch_holidays(year))
    # Filter op de periode
    return {d: name for d, name in holidays.items() if start <= d <= end}


if __name__ == "__main__":
    # Test
    from datetime import date
    year = 2025
    holidays = get_dutch_holidays(year)
    print(f"Nederlandse feestdagen {year}:")
    for d, name in sorted(holidays.items()):
        print(f"  {d.strftime('%d-%m-%Y')} ({d.strftime('%A')}): {name}")
