"""
Szuka potencjalnych klientów dla agencji UI/reklamowej:
aplikacje z App Store, które mają mało ocen (proxy dla małej liczby
użytkowników).

Wykorzystuje publiczne iTunes Search API (bez klucza, bez limitu poza
rozsądnym throttlingiem). Nie daje dokładnej liczby instalacji ani nie
ocenia estetyki automatycznie - to tylko pierwsze sito. Finalną ocenę
"czy jest brzydkie" i tak trzeba zrobić ręcznie, patrząc na screeny
i opis w App Store.

Użycie:
    pip install requests
    python find_app_leads.py
"""

import csv
import time
from datetime import datetime, timezone
from typing import Optional

import requests

# --- KONFIGURACJA -----------------------------------------------------

# Frazy do przeszukania - dopisz swoje nisze. Im węższa/bardziej lokalna
# fraza, tym większa szansa trafić na appki robione "na szybko" przez
# freelancera i potem zapomniane.
TERMS = [
    "salon fryzjerski rezerwacja",
    "gabinet kosmetyczny rezerwacja",
    "warsztat samochodowy",
    "restauracja zamówienia",
    "pizza zamówienie dostawa",
    "siłownia grafik zajęć",
    "gabinet stomatologiczny wizyta",
    "klub fitness karnet",
    "kwiaciarnia dostawa",
    "myjnia samochodowa",
    "wypożyczalnia sprzętu",
    "sklep lokalny zamówienia",
    "przychodnia rejestracja",
    "hotel rezerwacja lokalny",
    "biuro rachunkowe klient",
]

COUNTRY = "pl"          # kraj sklepu App Store
MAX_RATING_COUNT = 50   # próg "mało użytkowników"
RESULTS_PER_TERM = 50

OUTPUT_CSV = "app_leads.csv"

# --- LOGIKA ------------------------------------------------------------

def search_apps(term: str, country: str = COUNTRY, limit: int = RESULTS_PER_TERM):
    url = "https://itunes.apple.com/search"
    params = {
        "term": term,
        "country": country,
        "entity": "software",
        "limit": limit,
    }
    r = requests.get(url, params=params, timeout=15)
    r.raise_for_status()
    return r.json().get("results", [])


def days_since_update(app: dict) -> Optional[int]:
    release_date_str = app.get("currentVersionReleaseDate")
    if not release_date_str:
        return None
    release_date = datetime.fromisoformat(release_date_str.replace("Z", "+00:00"))
    return (datetime.now(timezone.utc) - release_date).days


def is_candidate(app: dict) -> bool:
    rating_count = app.get("userRatingCount", 0) or 0
    return rating_count <= MAX_RATING_COUNT


def main():
    seen_ids = set()
    candidates = []

    for term in TERMS:
        print(f"Szukam: {term}")
        try:
            results = search_apps(term)
        except requests.RequestException as e:
            print(f"  błąd zapytania: {e}")
            continue

        for app in results:
            app_id = app.get("trackId")
            if app_id in seen_ids:
                continue
            seen_ids.add(app_id)

            if is_candidate(app):
                candidates.append({
                    "nazwa": app.get("trackName"),
                    "deweloper": app.get("sellerName"),
                    "liczba_ocen": app.get("userRatingCount", 0),
                    "ostatnia_aktualizacja": app.get("currentVersionReleaseDate", "")[:10],
                    "dni_od_aktualizacji": days_since_update(app),
                    "kategoria": app.get("primaryGenreName"),
                    "url": app.get("trackViewUrl"),
                    "znaleziono_dla_frazy": term,
                })

        time.sleep(0.5)  # nie spamuj API

    candidates.sort(key=lambda c: c["liczba_ocen"])

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=candidates[0].keys() if candidates else [
            "nazwa", "deweloper", "liczba_ocen", "ostatnia_aktualizacja",
            "dni_od_aktualizacji", "kategoria", "url", "znaleziono_dla_frazy"
        ])
        writer.writeheader()
        writer.writerows(candidates)

    print(f"\nZnaleziono {len(candidates)} kandydatów. Zapisano do {OUTPUT_CSV}")
    for c in candidates[:20]:
        print(f"- {c['nazwa']} ({c['deweloper']}) - {c['liczba_ocen']} ocen, "
              f"aktualizacja {c['ostatnia_aktualizacja']}")


if __name__ == "__main__":
    main()
