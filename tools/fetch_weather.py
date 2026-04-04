"""
Haalt historische uurlijkse weersdata op via de Open-Meteo Archive API.
Geen API key vereist. Data beschikbaar vanaf 1940.
"""

import requests
from datetime import date, datetime
from typing import Optional


GEOCODING_URL = "https://geocoding-api.open-meteo.com/v1/search"
ARCHIVE_URL = "https://archive-api.open-meteo.com/v1/archive"

# CAO Bouwnijverheid drempelwaarden — standaard bouw
WIND_THRESHOLD_MS = 13.9   # m/s = Beaufort 7
RAIN_THRESHOLD_MM = 2.0    # mm per uur
FROST_THRESHOLD_C = -3.0   # graden Celsius
UNWORKABLE_HOURS = 5       # minimaal aantal uren onwerkbaar per dag
WORKDAY_START = 7          # 07:00
WORKDAY_END = 17           # tot 17:00 (uren 7 t/m 16 = 10 uur)

# Drempelwaarden per werksoort
NORMEN = {
    "standaard": {
        "wind_ms": 13.9,   # Beaufort 7
        "regen_mm": 2.0,
        "vorst_c": -3.0,
        "wind_label": "Bft 7",
    },
    "hijswerk": {
        "wind_ms": 10.8,   # Beaufort 6
        "regen_mm": 2.0,
        "vorst_c": -3.0,
        "wind_label": "Bft 6",
    },
}


def get_coordinates(location: str) -> Optional[tuple[float, float, str]]:
    """
    Zoekt de coördinaten op voor een locatie in Nederland.
    Geeft (latitude, longitude, display_name) terug, of None als niet gevonden.
    """
    params = {
        "name": location,
        "country": "NL",
        "count": 1,
        "language": "nl",
        "format": "json"
    }
    try:
        response = requests.get(GEOCODING_URL, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
        if not data.get("results"):
            return None
        result = data["results"][0]
        return result["latitude"], result["longitude"], result["name"]
    except requests.RequestException as e:
        raise ConnectionError(f"Fout bij ophalen locatie: {e}")


def fetch_hourly_weather(lat: float, lon: float, start: date, end: date) -> dict:
    """
    Haalt uurlijkse weersdata op voor de opgegeven coördinaten en periode.
    Geeft ruwe API response data terug.
    """
    params = {
        "latitude": lat,
        "longitude": lon,
        "start_date": start.isoformat(),
        "end_date": end.isoformat(),
        "hourly": "temperature_2m,precipitation,wind_speed_10m",
        "timezone": "Europe/Amsterdam",
        "wind_speed_unit": "ms"
    }
    try:
        response = requests.get(ARCHIVE_URL, params=params, timeout=30)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        raise ConnectionError(f"Fout bij ophalen weersdata: {e}")


def parse_weather_per_day(weather_data: dict, norm: str = "standaard") -> dict[date, dict]:
    """
    Verwerkt de uurlijkse API data naar een per-dag samenvatting.
    norm: 'standaard' of 'hijswerk'
    """
    drempel = NORMEN.get(norm, NORMEN["standaard"])
    wind_threshold = drempel["wind_ms"]
    regen_threshold = drempel["regen_mm"]
    vorst_threshold = drempel["vorst_c"]
    wind_label = drempel["wind_label"]

    hourly = weather_data.get("hourly", {})
    times = hourly.get("time", [])
    temperatures = hourly.get("temperature_2m", [])
    precipitations = hourly.get("precipitation", [])
    wind_speeds = hourly.get("wind_speed_10m", [])

    # Groepeer per dag
    daily = {}
    for i, time_str in enumerate(times):
        dt = datetime.fromisoformat(time_str)
        day = dt.date()
        hour = dt.hour

        # Alleen werkuren meenemen (07:00 t/m 16:00)
        if hour < WORKDAY_START or hour >= WORKDAY_END:
            continue

        if day not in daily:
            daily[day] = {
                "wind_hours": 0, "frost_hours": 0,
                "rain_hours_zwaar": 0,  # > 5 mm/u
                "rain_total": 0.0,
                "rain_max": 0.0,
                "wind_sum": 0.0, "wind_count": 0,
                "temp_sum": 0.0, "temp_count": 0,
            }

        temp = temperatures[i] if i < len(temperatures) else None
        precip = precipitations[i] if i < len(precipitations) else None
        wind = wind_speeds[i] if i < len(wind_speeds) else None

        if wind is not None:
            if wind >= wind_threshold:
                daily[day]["wind_hours"] += 1
            daily[day]["wind_sum"] += wind
            daily[day]["wind_count"] += 1
        if precip is not None:
            if precip > 5.0:
                daily[day]["rain_hours_zwaar"] += 1
            daily[day]["rain_total"] += precip
            if precip > daily[day]["rain_max"]:
                daily[day]["rain_max"] = precip
        if temp is not None:
            if temp <= vorst_threshold:
                daily[day]["frost_hours"] += 1
            daily[day]["temp_sum"] += temp
            daily[day]["temp_count"] += 1

    # Bepaal per dag of het onwerkbaar was
    result = {}
    for day, counts in daily.items():
        reasons = []

        # Regenlogica:
        # Onwerkbaar als totale neerslag > 20 mm/dag (werkuren)
        # EN/OF meer dan 5 mm/u gedurende 2 uur of meer
        regen_totaal = counts["rain_total"]
        uren_zwaar = counts["rain_hours_zwaar"]
        regen_onwerkbaar = regen_totaal > 20.0 or uren_zwaar >= 2
        regen_mogelijk = (not regen_onwerkbaar) and (10.0 < regen_totaal <= 20.0)
        if regen_onwerkbaar:
            redenen_regen = []
            if regen_totaal > 20.0:
                redenen_regen.append(f"totaal {round(regen_totaal, 1)}mm")
            if uren_zwaar >= 2:
                redenen_regen.append(f"{uren_zwaar}u >5mm/u")
            reasons.append(f"Neerslag ({', '.join(redenen_regen)})")

        if counts["wind_hours"] >= UNWORKABLE_HOURS:
            reasons.append(f"Wind ({counts['wind_hours']}u \u2265 {wind_label})")
        if counts["frost_hours"] >= UNWORKABLE_HOURS:
            reasons.append(f"Vorst ({counts['frost_hours']}u \u2264 -3\u00b0C)")

        wind_gem = (counts["wind_sum"] / counts["wind_count"]) if counts["wind_count"] else None
        temp_gem = (counts["temp_sum"] / counts["temp_count"]) if counts["temp_count"] else None

        result[day] = {
            "unworkable": len(reasons) > 0,
            "mogelijk_onwerkbaar": regen_mogelijk,
            "reasons": reasons,
            "wind_hours": counts["wind_hours"],
            "rain_hours_zwaar": counts["rain_hours_zwaar"],
            "rain_totaal_mm": round(regen_totaal, 1),
            "rain_max_mm": round(counts["rain_max"], 1),
            "frost_hours": counts["frost_hours"],
            "wind_gem_ms": round(wind_gem, 1) if wind_gem is not None else None,
            "temp_gem_c": round(temp_gem, 1) if temp_gem is not None else None,
        }

    return result


def get_weather_for_period(location: str, start: date, end: date, norm: str = "standaard") -> tuple[dict[date, dict], str]:
    """
    Combineert geocoding en weer ophalen.
    Geeft (weather_per_day, display_name) terug.
    """
    coords = get_coordinates(location)
    if coords is None:
        raise ValueError(f"Locatie '{location}' niet gevonden in Nederland.")

    lat, lon, display_name = coords
    raw_data = fetch_hourly_weather(lat, lon, start, end)
    weather_per_day = parse_weather_per_day(raw_data, norm=norm)

    return weather_per_day, display_name


if __name__ == "__main__":
    # Test
    from datetime import date
    start = date(2024, 1, 1)
    end = date(2024, 1, 31)
    weather, name = get_weather_for_period("Amsterdam", start, end)
    print(f"Weersdata voor {name}:")
    onwerkbaar = [(d, w) for d, w in sorted(weather.items()) if w["unworkable"]]
    print(f"Onwerkbare dagen door weer: {len(onwerkbaar)}")
    for d, w in onwerkbaar:
        print(f"  {d.strftime('%d-%m-%Y')}: {', '.join(w['reasons'])}")
