# Workflow: Werkbare Dagen Calculator

## Doel
Bereken het aantal werkbare en niet-werkbare dagen voor een bouwproject in een opgegeven periode en locatie in Nederland.

## Inputs
- Startdatum (dd-mm-yyyy)
- Einddatum (dd-mm-yyyy)
- Locatie (stad of gemeente in Nederland)

## Definitie Werkbare Dag
Een dag is **werkbaar** als aan alle volgende voorwaarden is voldaan:
1. Het is een weekdag (maandag t/m vrijdag)
2. Het is geen officiële Nederlandse feestdag
3. De weersomstandigheden maken de bouwplaats niet onwerkbaar (zie normen)

## CAO Bouwnijverheid — Normen Onwerkbaar Weer
Een dag is **onwerkbaar door weer** als gedurende minimaal 5 uur van de werkdag (07:00–17:00) één of meer van de volgende condities gelden:

| Conditie       | Drempelwaarde              |
|----------------|----------------------------|
| Wind           | ≥ 13.9 m/s (Beaufort 7)   |
| Neerslag       | ≥ 2 mm per uur             |
| Vorst          | ≤ -3°C                     |

## Officiële Nederlandse Feestdagen
- Nieuwjaarsdag (1 januari)
- Goede Vrijdag (variabel — vrijdag voor Pasen)
- Eerste Paasdag (variabel)
- Tweede Paasdag (variabel)
- Koningsdag (27 april, of 26 april als 27e een zondag is)
- Bevrijdingsdag (5 mei)
- Hemelvaartsdag (variabel — 39 dagen na Eerste Paasdag)
- Eerste Pinksterdag (variabel — 49 dagen na Eerste Paasdag)
- Tweede Pinksterdag (variabel — 50 dagen na Eerste Paasdag)
- Eerste Kerstdag (25 december)
- Tweede Kerstdag (26 december)

## Tools
- `tools/dutch_holidays.py` — berekent feestdagen voor een gegeven jaar
- `tools/fetch_weather.py` — haalt historische uurlijkse weersdata op via Open-Meteo API
- `tools/calculate_workdays.py` — combineert alle logica tot een resultaat
- `tools/gui.py` — de gebruikersinterface

## APIs
- **Geocoding**: `https://geocoding-api.open-meteo.com/v1/search?name={stad}&country=NL&count=1`
- **Historisch weer**: `https://archive-api.open-meteo.com/v1/archive`
  - Variabelen: `temperature_2m`, `precipitation`, `wind_speed_10m`
  - Tijdzone: `Europe/Amsterdam`
  - Interval: uurlijks

## Output
Per dag in de periode:
- Datum
- Dag van de week
- Status: Werkbaar / Weekend / Feestdag / Onwerkbaar weer
- Reden bij onwerkbaar (wind/regen/vorst + aantal uren)

Samenvatting:
- Totaal kalenderdagen
- Werkbare dagen
- Niet werkbare dagen (uitgesplitst: weekenden, feestdagen, onwerkbaar weer)

## Edge Cases
- Feestdag die op weekend valt → telt als weekend (niet dubbel tellen)
- Datum in de toekomst → Open-Meteo archive heeft geen toekomstige data; toon melding
- Locatie niet gevonden → vraag gebruiker om andere spelling
- API timeout → toon foutmelding in GUI, retry optie

## Bekende Beperkingen
- Open-Meteo archive data beschikbaar vanaf 1940
- Toekomstige datums kunnen niet worden geanalyseerd op weer
- Windsnelheid gemeten op 10m hoogte (standaard meteomast), op bouwplaats kan dit afwijken
