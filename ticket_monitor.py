#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ticket_monitor.py
=================

Dauerhafter Preis-Monitor fuer Event-Tickets, der ueber offizielle APIs
arbeitet und die Ergebnisse in eine Excel-Datei schreibt.

WARUM KEINE SCRAPER FUER VIAGOGO ODER STUBHUB
---------------------------------------------
Diese Seiten sind hinter Bot-Schutz (Cloudflare, Fingerprinting). Ein Scraper
dafuer funktioniert ein paar Tage, wird dann gesperrt, verstoesst gegen die
Nutzungsbedingungen und liefert unterwegs stillschweigend falsche Daten.
Dieses Programm nutzt stattdessen zwei offizielle, kostenlose Schnittstellen:

  1. SeatGeek Platform API
     SeatGeek IST ein Zweitmarkt. Die API gibt pro Event echte Listing-Preise
     zurueck: lowest_price, median_price, average_price, highest_price und
     listing_count. Dazu einen Popularitaets-Score zum Sortieren.
     Kostenloser Client ID: https://seatgeek.com/account/develop
     Abdeckung: sehr stark USA, schwach Europa.

  2. Ticketmaster Discovery API
     Liefert Erstverkaufspreise (priceRanges) als Nennwert-Referenz, dazu
     Vorverkaufsstart, Venue und Genre. Gute Abdeckung in Europa.
     Kostenloser API Key: https://developer.ticketmaster.com

EHRLICHE EINSCHRAENKUNG
-----------------------
Fuer europaeische Zweitmarktpreise gibt es keine kostenlose offizielle API.
SeatGeek deckt Europa kaum ab. Fuer Madrid oder Deutschland bekommst du mit
diesem Programm zuverlaessig die Erstverkaufsdaten von Ticketmaster und die
Zweitmarktdaten nur dort, wo SeatGeek Events fuehrt. Das ist keine Schwaeche
des Programms, sondern des Marktes.

BENUTZUNG
---------
    pip install requests openpyxl

    python ticket_monitor.py --demo            # Testlauf ohne API-Keys
    python ticket_monitor.py --once            # ein Durchlauf
    python ticket_monitor.py                   # Dauerbetrieb
    python ticket_monitor.py --export-only     # nur Excel neu bauen

Die Datenbank (SQLite) haelt die komplette Historie. Die Excel-Datei wird bei
jedem Durchlauf neu geschrieben und ist damit immer aktuell.
"""

import argparse
import json
import os
import random
import sqlite3
import sys
import time
from datetime import datetime, timedelta, timezone

try:
    import requests
except ImportError:
    requests = None

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# KONFIGURATION
# ============================================================

CONFIG = {
    # --- Zugangsdaten -------------------------------------------------
    # Entweder hier eintragen oder als Umgebungsvariablen setzen.
    "seatgeek_client_id": os.environ.get("SEATGEEK_CLIENT_ID", ""),
    "ticketmaster_api_key": os.environ.get("TICKETMASTER_API_KEY", ""),

    # --- Was beobachtet wird ------------------------------------------
    # Staedte fuer die SeatGeek-Suche nach populaeren Events.
    "seatgeek_cities": ["New York", "Los Angeles", "Chicago", "Miami"],
    # Freitext-Suchen, die zusaetzlich laufen (Kuenstler, die dich interessieren).
    "watch_queries": [],
    # Nur Events mit mindestens so vielen Listings aufnehmen. Filtert Rauschen.
    "min_listings": 10,
    # Wie viele Events pro Stadt maximal, sortiert nach Popularitaet.
    "events_per_city": 50,

    # Ticketmaster: Laendercodes fuer die Nennwert-Referenz.
    "ticketmaster_countries": ["ES", "DE"],
    "ticketmaster_classification": "Music",
    "ticketmaster_size": 100,

    # --- Wirtschaftliche Annahmen -------------------------------------
    # Verkaeufergebuehr des Zweitmarkts. Viagogo nennt keinen festen Satz,
    # in der Praxis 10 bis 15 Prozent. Quelle: Plattform-Hilfeseiten.
    "verkaeufergebuehr": 0.125,
    # Ab diesem Aufschlag gilt ein Event als interessant.
    "schwelle_aufschlag": 1.40,

    # --- Betrieb -------------------------------------------------------
    "intervall_minuten": 360,          # alle 6 Stunden
    "datenbank": "ticketdaten.sqlite",
    "excel_datei": "ticket_monitor.xlsx",
    # Fuer die Webseite: JSON, das das Dashboard liest. Leer = kein JSON.
    "json_datei": "docs/data.json",
    "hoeflichkeits_pause": 0.4,        # Sekunden zwischen API-Aufrufen
    "timeout": 20,
}

SEATGEEK_URL = "https://api.seatgeek.com/2/events"
TICKETMASTER_URL = "https://app.ticketmaster.com/discovery/v2/events.json"


# ============================================================
# DATENBANK
# ============================================================

SCHEMA = """
CREATE TABLE IF NOT EXISTS snapshots (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    erfasst_am      TEXT NOT NULL,
    quelle          TEXT NOT NULL,
    event_id        TEXT NOT NULL,
    titel           TEXT,
    typ             TEXT,
    event_datum     TEXT,
    venue           TEXT,
    stadt           TEXT,
    land            TEXT,
    kapazitaet      INTEGER,
    popularitaet    REAL,
    listings        INTEGER,
    preis_min       REAL,
    preis_median    REAL,
    preis_schnitt   REAL,
    preis_max       REAL,
    nennwert_min    REAL,
    nennwert_max    REAL,
    vvk_start       TEXT,
    waehrung        TEXT,
    url             TEXT
);
CREATE INDEX IF NOT EXISTS idx_event ON snapshots(event_id, erfasst_am);
CREATE INDEX IF NOT EXISTS idx_zeit  ON snapshots(erfasst_am);
"""


def db_oeffnen(pfad):
    verbindung = sqlite3.connect(pfad)
    verbindung.row_factory = sqlite3.Row
    verbindung.executescript(SCHEMA)
    verbindung.commit()
    return verbindung


def snapshot_schreiben(verbindung, zeilen):
    if not zeilen:
        return 0
    spalten = ["erfasst_am", "quelle", "event_id", "titel", "typ", "event_datum",
               "venue", "stadt", "land", "kapazitaet", "popularitaet", "listings",
               "preis_min", "preis_median", "preis_schnitt", "preis_max",
               "nennwert_min", "nennwert_max", "vvk_start", "waehrung", "url"]
    platzhalter = ",".join("?" * len(spalten))
    sql = "INSERT INTO snapshots (%s) VALUES (%s)" % (",".join(spalten), platzhalter)
    verbindung.executemany(sql, [[z.get(s) for s in spalten] for z in zeilen])
    verbindung.commit()
    return len(zeilen)


# ============================================================
# DATENQUELLEN
# ============================================================

def hole(url, parameter, timeout):
    """Ein GET mit Fehlerbehandlung. Gibt dict oder None zurueck."""
    if requests is None:
        raise RuntimeError("Das Paket 'requests' fehlt. Bitte: pip install requests")
    try:
        antwort = requests.get(url, params=parameter, timeout=timeout,
                               headers={"User-Agent": "ticket-monitor/1.0"})
    except requests.RequestException as fehler:
        print("  Netzwerkfehler: %s" % fehler)
        return None

    if antwort.status_code == 401 or antwort.status_code == 403:
        print("  Zugriff verweigert (%s). Pruefe den API-Schluessel." % antwort.status_code)
        return None
    if antwort.status_code == 429:
        print("  Rate Limit erreicht. Pausiere 60 Sekunden.")
        time.sleep(60)
        return None
    if antwort.status_code != 200:
        print("  HTTP %s von %s" % (antwort.status_code, url))
        return None
    try:
        return antwort.json()
    except ValueError:
        print("  Antwort war kein JSON.")
        return None


def seatgeek_abfragen(config, jetzt):
    """Populaere Events mit echten Zweitmarkt-Preisstatistiken."""
    client_id = config["seatgeek_client_id"]
    if not client_id:
        print("SeatGeek: kein Client ID gesetzt, wird uebersprungen.")
        return []

    gesammelt = {}
    anfragen = []
    for stadt in config["seatgeek_cities"]:
        anfragen.append({"venue.city": stadt})
    for begriff in config["watch_queries"]:
        anfragen.append({"q": begriff})

    for anfrage in anfragen:
        parameter = dict(anfrage)
        parameter.update({
            "client_id": client_id,
            "per_page": min(100, config["events_per_city"]),
            "sort": "score.desc",          # Popularitaet zuerst
            "datetime_utc.gte": jetzt.strftime("%Y-%m-%dT%H:%M:%S"),
        })
        etikett = anfrage.get("venue.city") or anfrage.get("q")
        print("SeatGeek: frage '%s' ab" % etikett)
        daten = hole(SEATGEEK_URL, parameter, config["timeout"])
        time.sleep(config["hoeflichkeits_pause"])
        if not daten:
            continue

        for ev in daten.get("events", []):
            stats = ev.get("stats") or {}
            listings = stats.get("listing_count")
            if listings is None or listings < config["min_listings"]:
                continue
            venue = ev.get("venue") or {}
            gesammelt[str(ev.get("id"))] = {
                "erfasst_am": jetzt.isoformat(timespec="seconds"),
                "quelle": "seatgeek",
                "event_id": str(ev.get("id")),
                "titel": ev.get("title"),
                "typ": ev.get("type"),
                "event_datum": (ev.get("datetime_local") or "")[:10],
                "venue": venue.get("name"),
                "stadt": venue.get("city"),
                "land": venue.get("country"),
                "kapazitaet": venue.get("capacity"),
                "popularitaet": ev.get("score"),
                "listings": listings,
                "preis_min": stats.get("lowest_price"),
                "preis_median": stats.get("median_price"),
                "preis_schnitt": stats.get("average_price"),
                "preis_max": stats.get("highest_price"),
                "nennwert_min": None,
                "nennwert_max": None,
                "vvk_start": (ev.get("announce_date") or "")[:10] or None,
                "waehrung": "USD",
                "url": ev.get("url"),
            }
    print("SeatGeek: %d Events mit Preisdaten" % len(gesammelt))
    return list(gesammelt.values())


def ticketmaster_backfill(config, jahre, verbindung):
    """
    Holt VERGANGENE Events von Ticketmaster, Monat fuer Monat rueckwaerts.

    Warum monatsweise: Die Discovery API deckelt die Tiefe der Blaetterung
    (size * page darf 1000 nicht ueberschreiten). Ein enges Zeitfenster pro
    Abfrage umgeht das und holt trotzdem alles.

    Was du dadurch bekommst: historische NENNWERTE, Venues, Termine und Genres.
    Was du dadurch NICHT bekommst: historische Zweitmarktpreise. Die
    veroeffentlicht keine Plattform rueckwirkend, weder Viagogo noch StubHub
    noch SeatGeek. Wer etwas anderes behauptet, verkauft dir geschaetzte Daten.

    Die Abdeckung wird schlechter, je weiter du zurueckgehst. Zwei bis drei
    Jahre sind realistisch, darueber hinaus wird es lueckenhaft.
    """
    schluessel = config["ticketmaster_api_key"]
    if not schluessel:
        print("Backfill braucht einen Ticketmaster API Key.")
        return 0

    ende = datetime.now(timezone.utc).replace(day=1)
    gesamt = 0
    monate = int(jahre * 12)

    for schritt in range(monate):
        fenster_ende = ende - timedelta(days=schritt * 30)
        fenster_start = fenster_ende - timedelta(days=30)

        for land in config["ticketmaster_countries"]:
            seite = 0
            while seite < 5:                       # max 500 Events pro Monat und Land
                parameter = {
                    "apikey": schluessel,
                    "countryCode": land,
                    "classificationName": config["ticketmaster_classification"],
                    "size": 100,
                    "page": seite,
                    "startDateTime": fenster_start.strftime("%Y-%m-%dT00:00:00Z"),
                    "endDateTime": fenster_ende.strftime("%Y-%m-%dT00:00:00Z"),
                }
                daten = hole(TICKETMASTER_URL, parameter, config["timeout"])
                time.sleep(config["hoeflichkeits_pause"])
                if not daten:
                    break

                events = (daten.get("_embedded") or {}).get("events", [])
                if not events:
                    break

                zeilen = []
                for ev in events:
                    spannen = ev.get("priceRanges") or []
                    if not spannen:
                        continue                   # ohne Preis ist die Zeile wertlos
                    orte = (ev.get("_embedded") or {}).get("venues") or [{}]
                    ort = orte[0]
                    verkauf = ((ev.get("sales") or {}).get("public") or {})
                    klass = (ev.get("classifications") or [{}])[0]
                    genre = ((klass.get("genre") or {}).get("name"))

                    zeilen.append({
                        "erfasst_am": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                        "quelle": "tm-historie",
                        "event_id": "hist-" + str(ev.get("id")),
                        "titel": ev.get("name"),
                        "typ": genre or "primary",
                        "event_datum": ((ev.get("dates") or {}).get("start") or {}).get("localDate"),
                        "venue": ort.get("name"),
                        "stadt": ((ort.get("city") or {}).get("name")),
                        "land": ((ort.get("country") or {}).get("countryCode")),
                        "kapazitaet": None,
                        "popularitaet": None,
                        "listings": None,
                        "preis_min": None,
                        "preis_median": None,
                        "preis_schnitt": None,
                        "preis_max": None,
                        "nennwert_min": min([s.get("min") for s in spannen if s.get("min") is not None], default=None),
                        "nennwert_max": max([s.get("max") for s in spannen if s.get("max") is not None], default=None),
                        "vvk_start": (verkauf.get("startDateTime") or "")[:10] or None,
                        "waehrung": spannen[0].get("currency"),
                        "url": ev.get("url"),
                    })

                gesamt += snapshot_schreiben(verbindung, zeilen)
                seite += 1
                if len(events) < 100:
                    break

        if schritt % 6 == 0:
            print("  Backfill bis %s, bisher %d Events" % (fenster_start.strftime("%Y-%m"), gesamt))

    print("Backfill abgeschlossen: %d historische Events gespeichert." % gesamt)
    return gesamt


# Recherchierte Vergleichswerte. Keine Schaetzungen, sondern belegte Faelle.
# Sie stehen als eigenes Excel-Blatt drin, damit du deine eigenen Messungen
# gegen dokumentierte Realitaet haeltst statt gegen Schlagzeilen.
BENCHMARKS = [
    ("Taylor Swift, Eras Tour 2022/23", "49 bis 499 USD", "Durchschnitt 2.424 USD",
     "ca. 5x bis 50x", "SeatGeek-Durchschnitt Nov 2022, Spitzen ueber 30.000 USD"),
    ("Taylor Swift, Eras, Plattformangaben", "49 bis 499 USD", "1.500 bis 3.000 USD",
     "bis 20x", "StubHub, SeatGeek, Vivid Seats"),
    ("Taylor Swift, Reputation Tour 2018", "k.A.", "Durchschnitt 279 USD",
     "Referenz", "derselbe Kuenstler vier Jahre frueher, zum Vergleich"),
    ("FTC-Verfahren, echte Broker-Bilanz", "744.970 USD Einkauf", "1.961.980 USD Verkauf",
     "2,63x brutto", "Gerichtsakte, industrieller Betrieb mit Bots, vor Gebuehren"),
    ("StubHub gesamt, 2024", "k.A.", "rund 218 USD je Ticket",
     "Marktdurchschnitt", "8,7 Mrd. USD GMS geteilt durch ueber 40 Mio. Tickets, S-1"),
    ("Oasis Reunion UK 2025", "k.A.", "ueber 50.000 Tickets storniert",
     "Totalverlust", "4 Prozent von 1,4 Mio. Tickets, zum Nennwert neu verteilt"),
]


def ticketmaster_abfragen(config, jetzt):
    """Erstverkaufspreise als Nennwert-Referenz, gute Europa-Abdeckung."""
    schluessel = config["ticketmaster_api_key"]
    if not schluessel:
        print("Ticketmaster: kein API Key gesetzt, wird uebersprungen.")
        return []

    gesammelt = {}
    for land in config["ticketmaster_countries"]:
        parameter = {
            "apikey": schluessel,
            "countryCode": land,
            "classificationName": config["ticketmaster_classification"],
            "size": config["ticketmaster_size"],
            "sort": "relevance,desc",
            "startDateTime": jetzt.strftime("%Y-%m-%dT%H:%M:%SZ"),
        }
        print("Ticketmaster: frage Land %s ab" % land)
        daten = hole(TICKETMASTER_URL, parameter, config["timeout"])
        time.sleep(config["hoeflichkeits_pause"])
        if not daten:
            continue

        events = (daten.get("_embedded") or {}).get("events", [])
        for ev in events:
            spannen = ev.get("priceRanges") or []
            nennwert_min = min([s.get("min") for s in spannen if s.get("min") is not None], default=None)
            nennwert_max = max([s.get("max") for s in spannen if s.get("max") is not None], default=None)
            waehrung = spannen[0].get("currency") if spannen else None

            orte = (ev.get("_embedded") or {}).get("venues") or [{}]
            ort = orte[0]
            verkauf = ((ev.get("sales") or {}).get("public") or {})

            gesammelt[str(ev.get("id"))] = {
                "erfasst_am": jetzt.isoformat(timespec="seconds"),
                "quelle": "ticketmaster",
                "event_id": str(ev.get("id")),
                "titel": ev.get("name"),
                "typ": "primary",
                "event_datum": ((ev.get("dates") or {}).get("start") or {}).get("localDate"),
                "venue": ort.get("name"),
                "stadt": ((ort.get("city") or {}).get("name")),
                "land": ((ort.get("country") or {}).get("countryCode")),
                "kapazitaet": None,
                "popularitaet": None,
                "listings": None,
                "preis_min": None,
                "preis_median": None,
                "preis_schnitt": None,
                "preis_max": None,
                "nennwert_min": nennwert_min,
                "nennwert_max": nennwert_max,
                "vvk_start": (verkauf.get("startDateTime") or "")[:10] or None,
                "waehrung": waehrung,
                "url": ev.get("url"),
            }
    print("Ticketmaster: %d Events" % len(gesammelt))
    return list(gesammelt.values())


def demo_daten(jetzt):
    """Synthetische Daten, damit du die Excel-Ausgabe ohne API-Keys siehst."""
    kuenstler = [
        ("Arctic Monkeys", "WiZink Center", "Madrid", 15000, 65),
        ("Rosalia", "Palau Sant Jordi", "Barcelona", 17000, 55),
        ("The Killers", "Barclays Arena", "Hamburg", 16000, 78),
        ("Fred again", "Razzmatazz", "Barcelona", 2500, 45),
        ("Bad Bunny", "Estadio Metropolitano", "Madrid", 68000, 95),
        ("Jungle", "La Riviera", "Madrid", 2400, 38),
        ("Tame Impala", "Lanxess Arena", "Koeln", 18500, 72),
        ("Peggy Gou", "Fabrik", "Madrid", 3000, 42),
    ]
    zeilen = []
    for i, (name, venue, stadt, kapazitaet, nennwert) in enumerate(kuenstler):
        knappheit = 1.0 + (3.2 * (1.0 - min(kapazitaet, 40000) / 40000.0)) * random.uniform(0.6, 1.3)
        tief = round(nennwert * knappheit, 2)
        zeilen.append({
            "erfasst_am": jetzt.isoformat(timespec="seconds"),
            "quelle": "demo",
            "event_id": "demo-%d" % i,
            "titel": name,
            "typ": "concert",
            "event_datum": (jetzt + timedelta(days=random.randint(40, 300))).strftime("%Y-%m-%d"),
            "venue": venue,
            "stadt": stadt,
            "land": "ES" if stadt in ("Madrid", "Barcelona") else "DE",
            "kapazitaet": kapazitaet,
            "popularitaet": round(random.uniform(0.55, 0.95), 3),
            "listings": random.randint(15, 900),
            "preis_min": tief,
            "preis_median": round(tief * random.uniform(1.2, 1.7), 2),
            "preis_schnitt": round(tief * random.uniform(1.3, 1.9), 2),
            "preis_max": round(tief * random.uniform(2.5, 6.0), 2),
            "nennwert_min": float(nennwert),
            "nennwert_max": float(nennwert) * 2.2,
            "vvk_start": (jetzt - timedelta(days=random.randint(20, 120))).strftime("%Y-%m-%d"),
            "waehrung": "EUR",
            "url": "https://example.invalid/%d" % i,
        })
    print("Demo: %d synthetische Events erzeugt" % len(zeilen))
    return zeilen


# ============================================================
# AUSWERTUNG
# ============================================================

def aktuelle_lage(verbindung):
    """Pro Event der juengste Snapshot, angereichert mit Historie."""
    sql = """
    SELECT s.* FROM snapshots s
    JOIN (SELECT event_id, MAX(erfasst_am) AS neuest
          FROM snapshots GROUP BY event_id) letzte
      ON s.event_id = letzte.event_id AND s.erfasst_am = letzte.neuest
    ORDER BY s.popularitaet DESC NULLS LAST, s.listings DESC
    """
    try:
        zeilen = [dict(r) for r in verbindung.execute(sql)]
    except sqlite3.OperationalError:
        # aeltere SQLite-Versionen kennen NULLS LAST nicht
        zeilen = [dict(r) for r in verbindung.execute(sql.replace(" NULLS LAST", ""))]

    for zeile in zeilen:
        verlauf = verbindung.execute(
            "SELECT erfasst_am, preis_min FROM snapshots "
            "WHERE event_id = ? AND preis_min IS NOT NULL ORDER BY erfasst_am",
            (zeile["event_id"],)).fetchall()
        zeile["beobachtungen"] = len(verlauf)
        if len(verlauf) >= 2 and verlauf[0]["preis_min"]:
            zeile["trend"] = zeile["preis_min"] / verlauf[0]["preis_min"] - 1.0
        else:
            zeile["trend"] = None
    return zeilen


def nennwert_zuordnen(zeilen):
    """
    SeatGeek kennt keinen Nennwert, Ticketmaster keinen Zweitmarktpreis.
    Ueber eine normalisierte Titelform werden beide Seiten verknuepft,
    wo dasselbe Event in beiden Quellen auftaucht.
    """
    def schluessel(t):
        if not t:
            return ""
        t = t.lower()
        for weg in [" tickets", " tour", " live", " concert", ":", "|", "-", "'"]:
            t = t.replace(weg, " ")
        return " ".join(t.split())[:40]

    nennwerte = {}
    for z in zeilen:
        if z.get("nennwert_min"):
            nennwerte[schluessel(z.get("titel"))] = (z["nennwert_min"], z.get("waehrung"))

    getroffen = 0
    for z in zeilen:
        if z.get("nennwert_min") or not z.get("preis_min"):
            continue
        treffer = nennwerte.get(schluessel(z.get("titel")))
        if treffer:
            z["nennwert_min"] = treffer[0]
            z["nennwert_quelle"] = "aus Ticketmaster zugeordnet"
            getroffen += 1
    if getroffen:
        print("Nennwert-Zuordnung: %d Events verknuepft" % getroffen)
    return zeilen


# ============================================================
# EXCEL
# ============================================================

SCHRIFT = "Arial"
KOPF_FUELLUNG = PatternFill("solid", fgColor="1F3243")
BAND = PatternFill("solid", fgColor="F2F2F2")
GELB = PatternFill("solid", fgColor="FFF2CC")
RAHMEN = Border(bottom=Side(style="thin", color="D0D0D0"))


def kopfzeile_schreiben(blatt, ueberschriften, zeile=1):
    for spalte, text in enumerate(ueberschriften, start=1):
        zelle = blatt.cell(row=zeile, column=spalte, value=text)
        zelle.font = Font(name=SCHRIFT, bold=True, color="FFFFFF", size=10)
        zelle.fill = KOPF_FUELLUNG
        zelle.alignment = Alignment(vertical="center", wrap_text=True)
    blatt.freeze_panes = blatt.cell(row=zeile + 1, column=1)
    blatt.row_dimensions[zeile].height = 30


def breiten_setzen(blatt, breiten):
    for i, b in enumerate(breiten, start=1):
        blatt.column_dimensions[get_column_letter(i)].width = b


def excel_bauen(zeilen, config, pfad, zeilen_alle_fuer_saison=None):
    if zeilen_alle_fuer_saison is None:
        zeilen_alle_fuer_saison = zeilen
    mappe = Workbook()

    # ---------- Blatt 1: Auswertung ----------
    blatt = mappe.active
    blatt.title = "Auswertung"

    ueberschriften = [
        "Event", "Typ", "Datum", "Venue", "Stadt", "Land", "Quelle",
        "Popularitaet", "Listings", "Nennwert", "Tiefster Preis", "Median",
        "Durchschnitt", "Hoechster Preis", "Waehrung",
        "Aufschlag", "Break-even", "Netto nach Gebuehr", "Marge",
        "Tage bis Event", "Rendite p.a.", "Trend seit Start", "Beobachtungen", "Urteil"
    ]
    kopfzeile_schreiben(blatt, ueberschriften)
    breiten_setzen(blatt, [30, 10, 11, 24, 14, 7, 12, 12, 9, 11, 13, 11, 12, 14, 9,
                           10, 11, 16, 11, 13, 11, 14, 13, 14])

    heute = datetime.now(timezone.utc).date()
    gebuehr_zelle = "Konfiguration!$B$4"
    schwelle_zelle = "Konfiguration!$B$5"

    for i, z in enumerate(zeilen):
        r = i + 2
        blatt.cell(row=r, column=1, value=z.get("titel"))
        blatt.cell(row=r, column=2, value=z.get("typ"))
        blatt.cell(row=r, column=3, value=z.get("event_datum"))
        blatt.cell(row=r, column=4, value=z.get("venue"))
        blatt.cell(row=r, column=5, value=z.get("stadt"))
        blatt.cell(row=r, column=6, value=z.get("land"))
        blatt.cell(row=r, column=7, value=z.get("quelle"))
        blatt.cell(row=r, column=8, value=z.get("popularitaet"))
        blatt.cell(row=r, column=9, value=z.get("listings"))
        blatt.cell(row=r, column=10, value=z.get("nennwert_min"))
        blatt.cell(row=r, column=11, value=z.get("preis_min"))
        blatt.cell(row=r, column=12, value=z.get("preis_median"))
        blatt.cell(row=r, column=13, value=z.get("preis_schnitt"))
        blatt.cell(row=r, column=14, value=z.get("preis_max"))
        blatt.cell(row=r, column=15, value=z.get("waehrung"))

        # Ab hier Formeln, damit die Tabelle neu rechnet, wenn du die
        # Gebuehr im Blatt "Konfiguration" aenderst.
        blatt.cell(row=r, column=16, value='=IFERROR(K%d/J%d,"")' % (r, r))
        blatt.cell(row=r, column=17, value='=IFERROR(J%d/(1-%s),"")' % (r, gebuehr_zelle))
        blatt.cell(row=r, column=18, value='=IFERROR(K%d*(1-%s),"")' % (r, gebuehr_zelle))
        blatt.cell(row=r, column=19, value='=IFERROR(R%d-J%d,"")' % (r, r))

        tage = None
        if z.get("event_datum"):
            try:
                tage = (datetime.strptime(z["event_datum"], "%Y-%m-%d").date() - heute).days
            except ValueError:
                tage = None
        blatt.cell(row=r, column=20, value=tage)
        blatt.cell(row=r, column=21, value='=IFERROR((S%d/J%d)*(365/T%d),"")' % (r, r, r))
        blatt.cell(row=r, column=22, value=z.get("trend"))
        blatt.cell(row=r, column=23, value=z.get("beobachtungen"))
        blatt.cell(row=r, column=24,
                   value='=IF(J%d="","Nennwert fehlt",IF(P%d>=%s,"pruefen","zu niedrig"))'
                         % (r, r, schwelle_zelle))

        for spalte in range(1, len(ueberschriften) + 1):
            zelle = blatt.cell(row=r, column=spalte)
            zelle.font = Font(name=SCHRIFT, size=10)
            zelle.border = RAHMEN
            if i % 2 == 1:
                zelle.fill = BAND

        for spalte in (10, 11, 12, 13, 14, 17, 18, 19):
            blatt.cell(row=r, column=spalte).number_format = '#,##0.00;(#,##0.00);-'
        blatt.cell(row=r, column=16).number_format = '0.00"x"'
        blatt.cell(row=r, column=21).number_format = '0.0%;(0.0%);-'
        blatt.cell(row=r, column=22).number_format = '+0.0%;-0.0%;-'
        blatt.cell(row=r, column=8).number_format = '0.000'

    blatt.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(ueberschriften)), max(2, len(zeilen) + 1))

    # ---------- Blatt 2: Verlauf ----------
    verlauf = mappe.create_sheet("Verlauf")
    v_kopf = ["Erfasst am", "Quelle", "Event-ID", "Event", "Datum", "Stadt",
              "Listings", "Tiefster Preis", "Median", "Durchschnitt", "Hoechster Preis", "Waehrung"]
    kopfzeile_schreiben(verlauf, v_kopf)
    breiten_setzen(verlauf, [20, 12, 14, 30, 12, 14, 9, 13, 11, 13, 14, 9])

    # ---------- Blatt 3: Saisonalitaet ----------
    saison = mappe.create_sheet("Saisonalitaet")
    kopfzeile_schreiben(saison, ["Monat", "Anzahl Events", "Median Nennwert",
                                 "Tiefster Nennwert", "Hoechster Nennwert", "Anteil am Jahr"])
    breiten_setzen(saison, [14, 15, 17, 18, 19, 14])

    monatsnamen = ["Januar", "Februar", "Maerz", "April", "Mai", "Juni", "Juli",
                   "August", "September", "Oktober", "November", "Dezember"]
    eimer = {i: [] for i in range(1, 13)}
    for z in zeilen_alle_fuer_saison:
        datum = z.get("event_datum")
        nennwert = z.get("nennwert_min")
        if not datum or nennwert is None:
            continue
        try:
            monat = int(str(datum)[5:7])
        except (ValueError, IndexError):
            continue
        if 1 <= monat <= 12:
            eimer[monat].append(float(nennwert))

    gesamt_events = sum(len(v) for v in eimer.values())
    for monat in range(1, 13):
        werte = sorted(eimer[monat])
        r = monat + 1
        saison.cell(row=r, column=1, value=monatsnamen[monat - 1])
        saison.cell(row=r, column=2, value=len(werte))
        if werte:
            mitte = len(werte) // 2
            med = werte[mitte] if len(werte) % 2 else (werte[mitte - 1] + werte[mitte]) / 2
            saison.cell(row=r, column=3, value=round(med, 2))
            saison.cell(row=r, column=4, value=werte[0])
            saison.cell(row=r, column=5, value=werte[-1])
        saison.cell(row=r, column=6,
                    value='=IFERROR(B%d/SUM($B$2:$B$13),"")' % r)
        for spalte in range(1, 7):
            zelle = saison.cell(row=r, column=spalte)
            zelle.font = Font(name=SCHRIFT, size=10)
            zelle.border = RAHMEN
        for spalte in (3, 4, 5):
            saison.cell(row=r, column=spalte).number_format = '#,##0.00;(#,##0.00);-'
        saison.cell(row=r, column=6).number_format = '0.0%;(0.0%);-'

    saison.cell(row=15, column=1,
                value="Grundlage: %d Events mit Nennwert aus der Datenbank." % gesamt_events
                ).font = Font(name=SCHRIFT, size=9, color="404040")
    saison.cell(row=16, column=1,
                value="Live Nation gibt in den Pflichtangaben an, dass Q2 und Q3 die "
                      "ertragsstaerksten Quartale sind, weil Open-Air-Venues und Festivals "
                      "hauptsaechlich von Mai bis Oktober bespielt werden."
                ).font = Font(name=SCHRIFT, size=9, color="404040")
    saison.cell(row=17, column=1,
                value="Fuer dich heisst das: gekauft wird im Winter, gespielt wird im Sommer, "
                      "ausgezahlt wird nach dem Event. Dein Kapital ist also strukturell lange gebunden."
                ).font = Font(name=SCHRIFT, size=9, color="C00000")

    # ---------- Blatt 4: Vergleichswerte ----------
    bench = mappe.create_sheet("Vergleichswerte")
    kopfzeile_schreiben(bench, ["Fall", "Nennwert", "Zweitmarkt", "Faktor", "Quelle und Anmerkung"])
    breiten_setzen(bench, [38, 22, 26, 16, 62])
    for i, eintrag in enumerate(BENCHMARKS):
        r = i + 2
        for spalte, wert in enumerate(eintrag, start=1):
            zelle = bench.cell(row=r, column=spalte, value=wert)
            zelle.font = Font(name=SCHRIFT, size=9)
            zelle.alignment = Alignment(wrap_text=True, vertical="top")
            zelle.border = RAHMEN
        bench.row_dimensions[r].height = 28
    bench.cell(row=len(BENCHMARKS) + 3, column=1,
               value="Die aussagekraeftigste Zeile ist die Broker-Bilanz aus dem FTC-Verfahren: "
                     "2,63x brutto, vor Gebuehren, bei einem industriellen Betrieb mit Bots. "
                     "Nicht 20x. Miss dich daran, nicht an Schlagzeilen."
               ).font = Font(name=SCHRIFT, size=9, bold=True, color="C00000")

    # ---------- Blatt 5: Konfiguration ----------
    konf = mappe.create_sheet("Konfiguration")
    breiten_setzen(konf, [34, 18, 62])
    konf["A1"] = "Konfiguration und Annahmen"
    konf["A1"].font = Font(name=SCHRIFT, bold=True, size=13)
    konf["A3"] = "Parameter"; konf["B3"] = "Wert"; konf["C3"] = "Herkunft und Bedeutung"
    for spalte in ("A3", "B3", "C3"):
        konf[spalte].font = Font(name=SCHRIFT, bold=True, color="FFFFFF", size=10)
        konf[spalte].fill = KOPF_FUELLUNG

    eintraege = [
        ("Verkaeufergebuehr", config["verkaeufergebuehr"],
         "Zweitmarkt-Verkaeufergebuehr. Viagogo nennt keinen festen Satz, in der Praxis 10 bis 15 Prozent. Hier aenderbar, die Auswertung rechnet neu."),
        ("Schwelle Aufschlag", config["schwelle_aufschlag"],
         "Ab diesem Vielfachen des Nennwerts markiert die Spalte Urteil ein Event als pruefenswert."),
        ("Mindestanzahl Listings", config["min_listings"],
         "Events mit weniger Listings werden ignoriert, weil einzelne Preise dort nicht aussagekraeftig sind."),
        ("Intervall in Minuten", config["intervall_minuten"],
         "Abstand zwischen zwei Durchlaeufen im Dauerbetrieb."),
        ("Zuletzt aktualisiert", datetime.now().strftime("%Y-%m-%d %H:%M"),
         "Zeitpunkt des letzten Excel-Exports."),
    ]
    for i, (name, wert, erklaerung) in enumerate(eintraege):
        r = i + 4
        konf.cell(row=r, column=1, value=name).font = Font(name=SCHRIFT, size=10, bold=True)
        zelle = konf.cell(row=r, column=2, value=wert)
        zelle.font = Font(name=SCHRIFT, size=10, color="0000FF")
        zelle.fill = GELB
        konf.cell(row=r, column=3, value=erklaerung).font = Font(name=SCHRIFT, size=9)
        konf.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")
        konf.row_dimensions[r].height = 30
    konf["B4"].number_format = "0.0%"

    hinweise = [
        "",
        "Blau hinterlegte Zellen sind Eingaben. Alles andere rechnet sich daraus.",
        "",
        "Datenquellen:",
        "SeatGeek Platform API. SeatGeek ist selbst ein Zweitmarkt, die Preise sind echte Listing-Preise.",
        "Ticketmaster Discovery API. Liefert Erstverkaufspreise als Nennwert-Referenz.",
        "",
        "Was diese Datei NICHT enthaelt:",
        "Verkaufspreise. Kein Zweitmarkt veroeffentlicht, zu welchem Preis tatsaechlich verkauft wurde.",
        "Alle Preise hier sind Angebotspreise, also das, was Verkaeufer verlangen, nicht was gezahlt wurde.",
        "Der tatsaechliche Verkaufspreis liegt in der Regel darunter.",
        "",
        "Formeln in der Auswertung:",
        "Aufschlag = Tiefster Preis geteilt durch Nennwert",
        "Break-even = Nennwert geteilt durch (1 minus Gebuehr)",
        "Netto = Tiefster Preis mal (1 minus Gebuehr)",
        "Marge = Netto minus Nennwert",
        "Rendite p.a. = (Marge geteilt durch Nennwert) mal (365 geteilt durch Tage bis Event)",
    ]
    for i, text in enumerate(hinweise):
        r = 10 + i
        konf.cell(row=r, column=1, value=text).font = Font(
            name=SCHRIFT, size=9,
            bold=text.endswith(":"),
            color="C00000" if "NICHT" in text else "404040")

    mappe.save(pfad)
    return pfad


def verlauf_fuellen(verbindung, pfad, grenze=20000):
    """Historie nachtragen, in einem zweiten Durchgang wegen Speicherbedarf."""
    from openpyxl import load_workbook
    mappe = load_workbook(pfad)
    blatt = mappe["Verlauf"]
    sql = ("SELECT erfasst_am, quelle, event_id, titel, event_datum, stadt, listings, "
           "preis_min, preis_median, preis_schnitt, preis_max, waehrung "
           "FROM snapshots ORDER BY erfasst_am DESC LIMIT ?")
    for i, zeile in enumerate(verbindung.execute(sql, (grenze,))):
        r = i + 2
        for spalte, wert in enumerate(zeile, start=1):
            zelle = blatt.cell(row=r, column=spalte, value=wert)
            zelle.font = Font(name=SCHRIFT, size=9)
        for spalte in (8, 9, 10, 11):
            blatt.cell(row=r, column=spalte).number_format = '#,##0.00;(#,##0.00);-'
    mappe.save(pfad)



# ============================================================
# JSON-EXPORT FUER DAS WEB-DASHBOARD
# ============================================================

def json_export(zeilen, saison_basis, config, verbindung, pfad):
    """
    Schreibt die Daten als data.json fuer die statische Webseite.
    Bewusst klein gehalten: nur was das Dashboard wirklich anzeigt,
    plus die Preishistorie pro Event fuer die Verlaufskurven.
    """
    gebuehr = config["verkaeufergebuehr"]
    heute = datetime.now(timezone.utc).date()

    events = []
    for z in zeilen:
        verlauf = verbindung.execute(
            "SELECT erfasst_am, preis_min FROM snapshots "
            "WHERE event_id = ? AND preis_min IS NOT NULL "
            "ORDER BY erfasst_am", (z["event_id"],)).fetchall()

        nennwert = z.get("nennwert_min")
        preis = z.get("preis_min")
        tage = None
        if z.get("event_datum"):
            try:
                tage = (datetime.strptime(z["event_datum"], "%Y-%m-%d").date() - heute).days
            except ValueError:
                tage = None

        netto = preis * (1 - gebuehr) if preis else None
        marge = (netto - nennwert) if (netto is not None and nennwert) else None
        rendite = None
        if marge is not None and nennwert and tage and tage > 0:
            rendite = (marge / nennwert) * (365.0 / tage)

        events.append({
            "id": z["event_id"],
            "titel": z.get("titel"),
            "typ": z.get("typ"),
            "datum": z.get("event_datum"),
            "venue": z.get("venue"),
            "stadt": z.get("stadt"),
            "land": z.get("land"),
            "quelle": z.get("quelle"),
            "url": z.get("url"),
            "popularitaet": z.get("popularitaet"),
            "listings": z.get("listings"),
            "nennwert": nennwert,
            "preis_min": preis,
            "preis_median": z.get("preis_median"),
            "preis_schnitt": z.get("preis_schnitt"),
            "preis_max": z.get("preis_max"),
            "waehrung": z.get("waehrung"),
            "aufschlag": (preis / nennwert) if (preis and nennwert) else None,
            "break_even": (nennwert / (1 - gebuehr)) if nennwert else None,
            "netto": netto,
            "marge": marge,
            "tage": tage,
            "rendite_pa": rendite,
            "trend": z.get("trend"),
            "beobachtungen": len(verlauf),
            "verlauf": [[r["erfasst_am"][:10], r["preis_min"]] for r in verlauf][-60:],
        })

    eimer = {i: [] for i in range(1, 13)}
    for z in saison_basis:
        datum, nennwert = z.get("event_datum"), z.get("nennwert_min")
        if not datum or nennwert is None:
            continue
        try:
            monat = int(str(datum)[5:7])
        except (ValueError, IndexError):
            continue
        if 1 <= monat <= 12:
            eimer[monat].append(float(nennwert))

    saison = []
    for monat in range(1, 13):
        werte = sorted(eimer[monat])
        med = None
        if werte:
            mitte = len(werte) // 2
            med = werte[mitte] if len(werte) % 2 else (werte[mitte - 1] + werte[mitte]) / 2
        saison.append({"monat": monat, "anzahl": len(werte),
                       "median": round(med, 2) if med is not None else None})

    gesamt_snapshots = verbindung.execute("SELECT COUNT(*) FROM snapshots").fetchone()[0]

    paket = {
        "aktualisiert": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "konfiguration": {
            "gebuehr": gebuehr,
            "schwelle_aufschlag": config["schwelle_aufschlag"],
            "min_listings": config["min_listings"],
        },
        "statistik": {
            "events": len(events),
            "snapshots_gesamt": gesamt_snapshots,
            "mit_nennwert": len([e for e in events if e["nennwert"]]),
        },
        "events": events,
        "saison": saison,
        "benchmarks": [
            {"fall": b[0], "nennwert": b[1], "zweitmarkt": b[2], "faktor": b[3], "quelle": b[4]}
            for b in BENCHMARKS
        ],
    }

    ordner = os.path.dirname(pfad)
    if ordner and not os.path.isdir(ordner):
        os.makedirs(ordner, exist_ok=True)
    with open(pfad, "w", encoding="utf-8") as datei:
        json.dump(paket, datei, ensure_ascii=False, separators=(",", ":"))
    print("JSON geschrieben: %s (%d Events, %.1f kB)"
          % (pfad, len(events), os.path.getsize(pfad) / 1024.0))
    return pfad


# ============================================================
# ABLAUF
# ============================================================

def durchlauf(config, verbindung, demo=False):
    jetzt = datetime.now(timezone.utc)
    print("\n=== Durchlauf %s ===" % jetzt.strftime("%Y-%m-%d %H:%M:%S UTC"))

    if demo:
        zeilen = demo_daten(jetzt)
    else:
        zeilen = []
        zeilen.extend(seatgeek_abfragen(config, jetzt))
        zeilen.extend(ticketmaster_abfragen(config, jetzt))

    if not zeilen:
        print("Keine Daten erhalten. Pruefe API-Schluessel und Konfiguration.")
    else:
        anzahl = snapshot_schreiben(verbindung, zeilen)
        print("%d Snapshots gespeichert." % anzahl)

    export(config, verbindung)


def export(config, verbindung):
    zeilen = nennwert_zuordnen(aktuelle_lage(verbindung))
    saison_basis = [dict(r) for r in verbindung.execute(
        "SELECT event_datum, nennwert_min FROM snapshots "
        "WHERE nennwert_min IS NOT NULL AND event_datum IS NOT NULL "
        "GROUP BY event_id")]
    pfad = None
    if config.get("excel_datei"):
        pfad = excel_bauen(zeilen, config, config["excel_datei"], saison_basis)
        verlauf_fuellen(verbindung, pfad)
    if config.get("json_datei"):
        json_export(zeilen, saison_basis, config, verbindung, config["json_datei"])

    mit_nennwert = [z for z in zeilen if z.get("nennwert_min") and z.get("preis_min")]
    interessant = [z for z in mit_nennwert
                   if z["preis_min"] / z["nennwert_min"] >= config["schwelle_aufschlag"]]
    if pfad:
        print("Excel geschrieben: %s" % os.path.abspath(pfad))
    print("  %d Events gesamt, %d mit Nennwert, %d ueber der Schwelle"
          % (len(zeilen), len(mit_nennwert), len(interessant)))
    for z in interessant[:10]:
        print("    %-34s %6.2f -> %7.2f  (%.2fx, %d Listings)"
              % ((z.get("titel") or "")[:34], z["nennwert_min"], z["preis_min"],
                 z["preis_min"] / z["nennwert_min"], z.get("listings") or 0))


def main():
    zerleger = argparse.ArgumentParser(description="Dauerhafter Ticketpreis-Monitor mit Excel-Ausgabe")
    zerleger.add_argument("--once", action="store_true", help="nur ein Durchlauf, dann beenden")
    zerleger.add_argument("--demo", action="store_true", help="synthetische Daten, ohne API-Schluessel")
    zerleger.add_argument("--export-only", action="store_true", help="nur Excel aus der Datenbank neu bauen")
    zerleger.add_argument("--backfill", type=float, metavar="JAHRE",
                          help="historische Ticketmaster-Events der letzten N Jahre laden (z.B. 2)")
    zerleger.add_argument("--intervall", type=int, help="Minuten zwischen Durchlaeufen")
    zerleger.add_argument("--db", help="Pfad zur SQLite-Datei")
    zerleger.add_argument("--excel", help="Pfad zur Excel-Datei")
    zerleger.add_argument("--json", help="Pfad zur data.json fuer das Web-Dashboard")
    zerleger.add_argument("--no-excel", action="store_true",
                          help="keine Excel-Datei schreiben (fuer den Serverbetrieb)")
    argumente = zerleger.parse_args()

    config = dict(CONFIG)
    if argumente.intervall:
        config["intervall_minuten"] = argumente.intervall
    if argumente.db:
        config["datenbank"] = argumente.db
    if argumente.excel:
        config["excel_datei"] = argumente.excel
    if argumente.json:
        config["json_datei"] = argumente.json
    if argumente.no_excel:
        config["excel_datei"] = None

    verbindung = db_oeffnen(config["datenbank"])

    if argumente.export_only:
        export(config, verbindung)
        return

    if argumente.backfill:
        print("Backfill ueber %.1f Jahre. Das dauert einige Minuten." % argumente.backfill)
        ticketmaster_backfill(config, argumente.backfill, verbindung)
        export(config, verbindung)
        return

    if not argumente.demo and not config["seatgeek_client_id"] and not config["ticketmaster_api_key"]:
        print("Kein API-Schluessel gesetzt.\n"
              "  SeatGeek (Zweitmarktpreise):   https://seatgeek.com/account/develop\n"
              "  Ticketmaster (Nennwerte):      https://developer.ticketmaster.com\n"
              "Danach als Umgebungsvariable setzen oder oben in CONFIG eintragen.\n"
              "Zum Ansehen der Ausgabe ohne Schluessel: python ticket_monitor.py --demo --once")
        sys.exit(1)

    if argumente.once:
        durchlauf(config, verbindung, demo=argumente.demo)
        return

    print("Dauerbetrieb, alle %d Minuten. Beenden mit Strg+C." % config["intervall_minuten"])
    try:
        while True:
            try:
                durchlauf(config, verbindung, demo=argumente.demo)
            except Exception as fehler:
                print("Durchlauf fehlgeschlagen: %s" % fehler)
            naechster = datetime.now() + timedelta(minutes=config["intervall_minuten"])
            print("Naechster Durchlauf: %s" % naechster.strftime("%Y-%m-%d %H:%M"))
            time.sleep(config["intervall_minuten"] * 60)
    except KeyboardInterrupt:
        print("\nBeendet. Die Datenbank bleibt erhalten.")


if __name__ == "__main__":
    main()
