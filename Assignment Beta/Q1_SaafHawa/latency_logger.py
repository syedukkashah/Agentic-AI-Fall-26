"""
SaafHawa tools + latency logger (Assignment Beta, Q1, Section 2.6).

The four agent tools:
    get_air_quality  - real PM2.5 / PM10 from OpenAQ v3 (free API key needed)
    get_weather      - temperature, humidity, wind from Open-Meteo (no key)
    heat_index       - NOAA heat index from temperature + humidity (local)
    kb_lookup        - keyword search over knowledge_base.json (local)

Run it once on each of two different days:
    PowerShell:  $env:OPENAQ_API_KEY = "your-key";  python latency_logger.py
    bash:        OPENAQ_API_KEY=your-key python latency_logger.py

Every timing is appended to logs/tool_latency_log.csv (this is the raw log you submit).
"""
import csv
import json
import os
import statistics
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import requests

HERE = Path(__file__).parent
LOG_FILE = HERE / "logs" / "tool_latency_log.csv"
KB_FILE = HERE / "data" / "knowledge_base.json"

KARACHI = (24.8607, 67.0011)          # lat, lon of central Karachi
OPENAQ_URL = "https://api.openaq.org/v3"
OPENMETEO_URL = "https://api.open-meteo.com/v1/forecast"
CALLS_PER_API = 20                    # the assignment asks for >= 20 per API in total
PAUSE_SECONDS = 1.0                   # be polite to the free APIs


# ---------------------------------------------------------------- tools
def get_with_retry(url, tries=3, **kwargs):
    """requests.get with a few retries (the local DNS sometimes fails to resolve api.openaq.org)."""
    for attempt in range(tries):
        try:
            r = requests.get(url, timeout=20, **kwargs)
            r.raise_for_status()
            return r
        except requests.ConnectionError:
            if attempt == tries - 1:
                raise
            time.sleep(2)


def find_openaq_station(api_key, lat=KARACHI[0], lon=KARACHI[1]):
    """Pick a Karachi station that reported PM2.5 in the last 2 days, preferring one that also has PM10.
    Returns (location_id, name, {sensor_id: parameter})."""
    r = get_with_retry(f"{OPENAQ_URL}/locations", headers={"X-API-Key": api_key},
                       params={"coordinates": f"{lat},{lon}", "radius": 25000, "limit": 100})
    two_days_ago = (datetime.utcnow() - timedelta(days=2)).isoformat()
    live = []
    for loc in r.json()["results"]:
        sensors = {s["id"]: s["parameter"]["name"] for s in loc["sensors"]}
        last = (loc.get("datetimeLast") or {}).get("utc", "")
        if "pm25" in sensors.values() and last >= two_days_ago:
            live.append(("pm10" in sensors.values(), last, loc["id"], loc["name"], sensors))
    if not live:
        raise RuntimeError("No live OpenAQ station with PM2.5 near Karachi")
    has_pm10, last, loc_id, name, sensors = max(live)      # PM10 first, then most recent
    return loc_id, name, sensors


def get_air_quality(api_key, location_id, sensors):
    """Latest PM2.5 and PM10 (ug/m3) for one OpenAQ station. PM10 is None if the station doesn't measure it."""
    r = requests.get(f"{OPENAQ_URL}/locations/{location_id}/latest",
                     headers={"X-API-Key": api_key}, timeout=20)
    r.raise_for_status()
    readings = {"pm25": None, "pm10": None}
    for m in r.json()["results"]:
        name = sensors.get(m["sensorsId"])
        if name in readings:
            readings[name] = m["value"]
            readings[f"{name}_time"] = m["datetime"]["utc"]
    return readings


def get_weather(lat=KARACHI[0], lon=KARACHI[1]):
    """Current temperature (C), relative humidity (%) and wind speed (km/h)."""
    r = requests.get(OPENMETEO_URL, params={
        "latitude": lat, "longitude": lon,
        "current": "temperature_2m,relative_humidity_2m,wind_speed_10m",
    }, timeout=20)
    r.raise_for_status()
    cur = r.json()["current"]
    return {"temp_c": cur["temperature_2m"], "humidity": cur["relative_humidity_2m"],
            "wind_kmh": cur["wind_speed_10m"]}


def heat_index(temp_c, humidity):
    """NOAA heat index (Rothfusz regression), returned in Celsius."""
    t = temp_c * 9 / 5 + 32                                  # formula works in Fahrenheit
    simple = 0.5 * (t + 61.0 + (t - 68.0) * 1.2 + humidity * 0.094)
    if (simple + t) / 2 < 80:                                # mild weather: simple formula is enough
        hi = simple
    else:
        hi = (-42.379 + 2.04901523 * t + 10.14333127 * humidity
              - 0.22475541 * t * humidity - 6.83783e-3 * t ** 2
              - 5.481717e-2 * humidity ** 2 + 1.22874e-3 * t ** 2 * humidity
              + 8.5282e-4 * t * humidity ** 2 - 1.99e-6 * t ** 2 * humidity ** 2)
    return round((hi - 32) * 5 / 9, 1)


def load_kb():
    return json.loads(KB_FILE.read_text(encoding="utf-8"))


def kb_lookup(query, kb, top_k=2):
    """Return the top_k advisories whose keywords overlap most with the query."""
    words = set(query.lower().replace("?", " ").replace(",", " ").split())
    scored = [(len(words & set(item["keywords"])), item) for item in kb]
    scored.sort(key=lambda pair: pair[0], reverse=True)
    return [item for score, item in scored[:top_k] if score > 0]


# ---------------------------------------------------------------- timing
def time_call(fn, *args):
    """Run fn(*args) and return (milliseconds, ok, error_message)."""
    start = time.perf_counter()
    try:
        fn(*args)
        ok, err = True, ""
    except Exception as e:                   # log failures too; they are real latency events
        ok, err = False, str(e)[:120]
    return (time.perf_counter() - start) * 1000, ok, err


def append_rows(rows):
    LOG_FILE.parent.mkdir(exist_ok=True)
    new_file = not LOG_FILE.exists()
    with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if new_file:
            writer.writerow(["timestamp", "date", "tool", "latency_ms", "ok", "error"])
        writer.writerows(rows)


def main():
    api_key = os.environ.get("OPENAQ_API_KEY", "")
    if not api_key:
        sys.exit("Set the OPENAQ_API_KEY environment variable first (free key at https://explore.openaq.org).")

    kb = load_kb()
    location_id, station, sensors = find_openaq_station(api_key)
    print(f"Using OpenAQ station {location_id} ({station}); sample reading: {get_air_quality(api_key, location_id, sensors)}")
    print(f"Sample weather: {get_weather()}")

    sample_queries = ["Is it safe to go for an evening jog by the sea today?",
                      "Aj din 12 bajy temperature kitna rise hoga q k bachon ki school 12 bajy off hojaye ga?"]
    rows = []
    for i in range(CALLS_PER_API):
        now = datetime.now()
        stamp, day = now.isoformat(timespec="seconds"), now.date().isoformat()
        for tool, fn, args in [
            ("air_quality", get_air_quality, (api_key, location_id, sensors)),
            ("weather", get_weather, ()),
            ("heat_index", heat_index, (34.0 + i % 5, 55 + i % 20)),
            ("kb_lookup", kb_lookup, (sample_queries[i % 2], kb)),
        ]:
            ms, ok, err = time_call(fn, *args)
            rows.append([stamp, day, tool, round(ms, 4), ok, err])
        print(f"round {i + 1}/{CALLS_PER_API} done")
        time.sleep(PAUSE_SECONDS)

    append_rows(rows)
    print(f"\nAppended {len(rows)} rows to {LOG_FILE.name}")
    for tool in ["air_quality", "weather", "heat_index", "kb_lookup"]:
        times = [r[3] for r in rows if r[2] == tool and r[4]]
        if times:
            print(f"  {tool:12s} median = {statistics.median(times):9.3f} ms  (n={len(times)})")


if __name__ == "__main__":
    main()
