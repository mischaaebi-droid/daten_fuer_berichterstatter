"""Bundesgerichtsurteile des aktuellen Tages herunterladen und auswerten.

Erwartete Dateien im gleichen Ordner wie dieses Skript:
  - list_federal_judges_party.xlsx
  - ext_script_decision_match.py
  - ext_script_beschwerden_match.py

Benötigte Pakete:
  pip install playwright beautifulsoup4 lxml pandas openpyxl
  playwright install chromium
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
import unicodedata
from datetime import datetime, timedelta
from urllib.parse import urljoin, unquote
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, sync_playwright, Error as PlaywrightError
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook


BASE_URL = "https://search.bger.ch/ext/eurospider/live/de/php/aza/http/index_aza.php"
RSS_URL = "http://relevancy.bger.ch/feeds/aza_de.rss"
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_HTML_DIR = SCRIPT_DIR / "aktuelle_urteile"
DEFAULT_JUDGES_FILE = SCRIPT_DIR / "list_federal_judges_party.xlsx"


LEGAL_AREAS = {
    "I1": "Grundrecht",
    "I2": "Bürgerrecht und Ausländerrecht",
    "I3": "Staatshaftung",
    "I4": "Öffentliches Dienstverhältnis",
    "I5": "Politische Rechte",
    "I6": "Rechtshilfe und Auslieferung",
    "I7": "Unterrichtswesen und Berufsausbildung",
    "I9": "Ökologisches Gleichgewicht",
    "I10": "Sicherheits- und Friedenspolitik",
    "I11": "Öffentliche Finanzen & Abgaberecht",
    "I12": "Raumplanung und öffentliches Baurecht",
    "I13": "Enteignung",
    "I15": "Gesundheitswesen & soziale Sicherheit",
    "I16": "Wirtschaft",
    "I18": "Schuldbetreibungs- und Konkursrecht",
    "I21": "Immaterialgüter-, Wettbewerbs- und Kartellrecht",
    "I30": "Alters- und Hinterlassenenversicherung",
    "I31": "Invalidenversicherung",
    "I32": "Ergänzungsleistungen",
    "I33": "Berufliche Vorsorge",
    "I34": "Krankenversicherung",
    "I35": "Unfallversicherung",
    "I36": "Militärversicherung",
    "I37": "Erwerbsersatzordnung",
    "I38": "Familienzulagen und kantonale Sozialversicherung",
    "I39": "Arbeitslosenversicherung",
    "I141": "Energie",
    "I142": "Strassenbau und Strassenverkehr",
    "I143": "Verkehr (ohne Strassenverkehr)",
    "I144": "Post- und Fernmeldeverkehr",
    "I145": "Mass media",
    "I171": "Zivilprozess",
    "I172": "Strafprozess",
    "I173": "Verwaltungsverfahren",
    "I174": "Zuständigkeitsfragen / verfassungsmässiger Richter",
    "I176": "Schiedsgerichtsbarkeit",
    "I177": "Aufsichtsbeschwerden",
    "I192": "Personenrecht",
    "I193": "Familienrecht",
    "I194": "Erbrecht",
    "I195": "Sachenrecht",
    "I196": "Register",
    "I201": "Obligationenrecht (allgemein)",
    "I202": "Vertragsrecht",
    "I203": "Gesellschaftsrecht",
    "I205": "Haftpflichtrecht",
    "I221": "Strafrecht (allgemein)",
    "I222": "Straftaten",
    "I223": "Verwaltungsstrafrecht",
    "I224": "Straf- und Massnahmenvollzug",
    "I": "keine Zuordnung",
}

# RSS-Metadaten der im aktuellen Lauf gespeicherten HTML-Dateien.
RSS_METADATA: dict[str, dict[str, str | None]] = {}


def log(message: str, log_file: Path) -> None:
    print(message)
    with log_file.open("a", encoding="utf-8") as handle:
        handle.write(message + "\n")


def current_swiss_date() -> str:
    """Aktuelles Datum in der für die Bundesgerichts-URL nötigen Form."""
    return datetime.now(ZoneInfo("Europe/Zurich")).strftime("%Y%m%d")


def http_fetch_with_retry(request: urllib.request.Request, *, timeout: int = 30, max_attempts: int = 3):
    """HTTP-Abruf mit schonendem Retry bei 429/503."""
    for attempt in range(1, max_attempts + 1):
        try:
            return urllib.request.urlopen(request, timeout=timeout)
        except urllib.error.HTTPError as exc:
            if exc.code not in (429, 503) or attempt == max_attempts:
                raise
            wait_seconds = 60
            print(
                f"[HTTP {exc.code}] BGer vorübergehend nicht verfügbar. "
                f"Warte {wait_seconds} Sekunden (Versuch {attempt}/{max_attempts}) ..."
            )
            time.sleep(wait_seconds)


def case_key(value):
    match = re.search(r"\b(\d+[A-Z]_\d+)[/-](\d{4})\b", unquote(str(value)))
    return f"{match[1]}/{match[2]}" if match else None


def existing_decisions(day, output_dir):
    path = output_dir / "data" / f"decisions-{day[:4]}.json"
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))["decisions"]
    if not isinstance(data, dict):
        raise ValueError(f"Ungültige Jahresdatei: {path}")
    return data


def complete(entry):
    judges = entry.get("judges", [])
    return (bool(judges) and entry.get("judge_count") == len(judges)
            and len(judges) in (1, 3, 5)
            and all(j.get("full_name") and j.get("party") for j in judges)
            and entry.get("richter_vollstaendig", True))


def rss_metadata():
    # RSS dient ausschliesslich zur nachträglichen Rechtsgebiet-Zuordnung.
    request = urllib.request.Request(RSS_URL, headers={"User-Agent": "Mozilla/5.0"})
    with http_fetch_with_retry(request) as response:
        root = ET.fromstring(response.read())
    ns = {"bgerlaw": "http://relevancy.bger.ch/feeds/bgerlaw.xsd"}
    result = {}
    for item in root.findall("./channel/item"):
        matiere = item.find("bgerlaw:matiere", ns)
        code = (matiere.get("id") or matiere.get("ID")) if matiere is not None else None
        if code not in LEGAL_AREAS:
            continue
        for dossier in item.findall("bgerlaw:dossiers/bgerlaw:dossier", ns):
            key = case_key(dossier.text)
            if key:
                result[key] = {"legal_area_code": code, "legal_area": LEGAL_AREAS[code]}
    return result


class AccessStopped(RuntimeError):
    pass


def goto_retry(page, url):
    for attempt in range(3):
        response = page.goto(url, wait_until="domcontentloaded", timeout=60_000)
        if response is None:
            raise RuntimeError(f"Keine HTTP-Antwort: {url}")
        if response.status in (429, 503) and attempt < 2:
            time.sleep(60)
            continue
        if response.status in (403, 429, 503):
            raise AccessStopped(f"HTTP {response.status}: {url}")
        if response.status >= 400:
            raise RuntimeError(f"HTTP {response.status}: {url}")
        return


def scrape_daily_decisions(day, output_dir, log_file, existing, metadata):
    output_dir.mkdir(parents=True, exist_ok=True)
    saved_files = []
    stopped = False
    errors = 0
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            goto_retry(page, f"{BASE_URL}?date={day}&lang=de&mode=news")
            locator = page.locator('a[href*="highlight_docid="]')
            try:
                locator.first.wait_for(state="attached", timeout=10_000)
            except PlaywrightTimeoutError:
                pass
            links = list(dict.fromkeys(urljoin(page.url, href) for href in
                (locator.nth(i).get_attribute("href") for i in range(locator.count())) if href))
            log(f"{len(links)} Urteilslinks für {day}", log_file)
            for number, url in enumerate(links, 1):
                key = case_key(url)
                if key and complete(existing.get(key, {})):
                    continue
                try:
                    time.sleep(15)
                    goto_retry(page, url)
                    page.locator("#highlight_content").wait_for(state="attached", timeout=20_000)
                    html = page.content()
                    soup = BeautifulSoup(html, "lxml")
                    tag = soup.find("b")
                    actual = case_key(tag.get_text(strip=True)) if tag else None
                    if not actual or not soup.select_one("#highlight_content"):
                        raise ValueError("Kein gültiger Urteilstext")
                    if complete(existing.get(actual, {})):
                        continue
                    safe = actual.replace("/", "-")
                    target = output_dir / f"urteil__{number}__{number}_{day[:4]}_{day[4:6]}_{day[6:]}__{safe}.html"
                    target.write_text(html, encoding="utf-8")
                    saved_files.append(target)
                    RSS_METADATA[target.name] = metadata.get(actual, {})
                    log(f"Gespeichert: {target.name}", log_file)
                except AccessStopped as exc:
                    log(f"Abruf eingestellt: {exc}", log_file)
                    stopped = True
                    errors += 1
                    break
                except Exception as exc:
                    errors += 1
                    log(f"Fehler beim Urteil {url}: {exc}", log_file)
        finally:
            browser.close()
    return saved_files, stopped, errors


TEXT_CORRECTIONS = {

    
    "IIe Cour de droit pénal MM. et Mme les Juges fédéraux Abrecht, Président, Koch et Hofmann.":
        "IIe Cour de droit pénal Composition MM. et Mme les Juges fédéraux Abrecht, Président, Koch et Hofmann.",
    "I. öffentlich-rechtliche Abteilung Bundesrichter Müller, präsidierendes Mitglied, Bundesrichter Chaix, Kneubühler,":
        "I. öffentlich-rechtliche Abteilung Besetzung Bundesrichter Müller, präsidierendes Mitglied, Bundesrichter Chaix, Kneubühler,",
    "Gerichtsschreiberin Frey Krieger. A._________, Beschwerdeführer,":
        "Gerichtsschreiberin Frey Krieger. Verfahrensbeteiligte A._________, Beschwerdeführer,",
    "IIe Cour de droit public Mmes et MM. les Juges fédéraux Aubry Girardin, Présidente, Donzallaz, Hänni, Ryter et Kradolfer.":
        "IIe Cour de droit public Composition Mmes et MM. les Juges fédéraux Aubry Girardin, Présidente, Donzallaz, Hänni, Ryter et Kradolfer.",
    "I. öffentlich-rechtliche Abteilung Bundesrichter Haag, Präsident, nebenamtliche Bundesrichterin Pont Veuthey, nebenamtlicher Bundesrichter Mecca,":
        "I. öffentlich-rechtliche Abteilung Besetzung Bundesrichter Haag, Präsident, nebenamtliche Bundesrichterin Pont Veuthey, nebenamtlicher Bundesrichter Mecca,",
    "I Mme la Juge fédérale Kiss, Juge présidant. Greffier : M. Widmer.":
        "I Composition Mme la Juge fédérale Kiss, Juge présidant. Greffier : M. Widmer.",
    "II. zivilrechtliche Abteilung Bundesrichter Bovey, Präsident, Bundesrichter Herrmann, Hartmann,":
        "II. zivilrechtliche Abteilung Besetzung Bundesrichter Bovey, Präsident, Bundesrichter Herrmann, Hartmann,",
    "II. strafrechtliche Abteilung Bundesrichter Hurni, als Einzelrichter, Gerichtsschreiber Stadler.":
        "II. strafrechtliche Abteilung Besetzung Bundesrichter Hurni, als Einzelrichter, Gerichtsschreiber Stadler.",
    "Gerichtsschreiber Stadler. A.________, vertreten durch Rechtsanwalt Gandi Calan,":
        "Gerichtsschreiber Stadler. Verfahrensbeteiligte A.________, vertreten durch Rechtsanwalt Gandi Calan,",
}

NAME_CORRECTIONS = {

    "Knebühler": "Kneubühler",
    "Kneubüher": "Kneubühler",
    "Donzalaz": "Donzallaz",
    "MM. le Juges fédéraux Bovey": "Bovey",
    "et MM. les Juges fédéraux Donzallaz": "Donzallaz",
    "Hänni D": "D. Hänni",
    "Hänni D.": "D. Hänni",
    
    
    "Petrik": "Petrik-Haltiner", "Hermann": "Herrmann", "Muschetti": "Muschietti",
    "van de Graaaf": "van de Graaf", "Van de Graaf": "van de Graaf",
    "Hoffmann": "Hofmann", "H ofmann": "Hofmann", "Von Felten": "von Felten",
    "de Rossa": "De Rossa", "Koc h": "Koch", "Kradofler": "Kradolfer",
    "Pont Venthey": "Pont Veuthey", "Kneuhühler": "Kneubühler",
    "Müller Th": "Müller", "et Hofmann": "Hofmann",
    "Präsident van de Graaf": "van de Graaf", "Präsident Abrecht": "Abrecht",
    "les Juges fédéraux Abrecht": "Abrecht", "MM. Juges fédéraux Abrecht": "Abrecht",
    "M. Mmes les Juges fédéraux Abrecht": "Abrecht", "M. les Juge fédéral Merz": "Merz",
    "M. Donzallaz": "Donzallaz",
}


def detect_language(text: str) -> str | None:
    if "Besetzung" in text:
        return "de"
    if "Composition" in text:
        return "fr"
    if "Composizione" in text:
        return "it"
    return None


def extract_panel_block(text: str, language: str | None) -> str | None:
    match = None
    if language == "de":
        match = re.search(
            r"(Bundesrichter(?:in)?\s+.+?,\s*(?:Präsident|Präsidentin),?)\s+"
            r"Besetzung\s+(.*?)\s+Verfahrensbeteiligte", text, re.DOTALL | re.IGNORECASE,
        )
        if not match:
            match = re.search(r"Besetzung\s+(.*?)\s+Verfahrensbeteiligte", text, re.DOTALL)
    elif language == "fr":
        match = re.search(
            r"Composition\s+(.*?)\s+(?:Participant(?:e)?s? à la procédure|\d+[A-Z]?_\d+/\d{4})",
            text, re.DOTALL,
        )
    elif language == "it":
        match = re.search(r"Composizione\s+(.*?)\s+Partecipanti al procedimento", text, re.DOTALL)

    if not match:
        return None
    groups = match.groups()
    block = groups[0] if len(groups) == 1 else groups[0] + " " + groups[1]
    return re.sub(r"\s+", " ", block).strip()


def extract_president(block: str, language: str) -> str | None:
    """Erkennt den Vorsitz nach derselben Logik wie das bewährte Aufbereitungs-Notebook."""
    if language == "fr":
        pattern = (
            r"(?:M\.\s+le\s+Juge\s+fédéral|"
            r"Mme\s+la\s+Juge\s+fédérale|"
            r"MM\.\s+les\s+Juges\s+fédéraux|"
            r"MM\.\s+et\s+Mmes\s+les\s+Juges\s+fédéraux|"
            r"MM\s+et\s+Mme\s+les\s+Juges\s+fédéraux|"
            r"Mmes\s+et\s+MM\s+les\s+Juges\s+fédéraux|"
            r"Mmes\s+et\s+M\.\s+les\s+Juges\s+fédéraux|"
            r"MM\.\s+et\s+Mme\s+les\s+Juges\s+fédéraux|"
            r"MM\.\s+et\s+Mme\s+et\s+les\s+Juges\s+fédéraux)"
            r"\s+(.+?),\s*"
            r"(?:Président|Présidente|Juge présidant|Juge présidante)"
        )
        flags = re.DOTALL | re.IGNORECASE
    elif language == "de":
        pattern = (
            r"(?:nebenamtliche[rn]?\s+)?Bundesrichter(?:in)?\s+(.+?),\s*"
            r"(?:Präsident|Präsidentin|(?:als\s+)?präsidierendes Mitglied)"
        )
        flags = 0
    elif language == "it":
        pattern = (
            r"Giudic[ei]\s+federal[ei]\s+(.+?),\s*"
            r"(?:Presidente|Giudice presidente)"
        )
        flags = re.IGNORECASE
    else:
        return None

    match = re.search(pattern, block, flags)
    return match.group(1).strip() if match else None


def extract_single_judge(block: str, language: str) -> str | None:
    """Separate Einzelrichter-Erkennung aus dem Vorbild-Notebook."""
    if language == "de":
        pattern = (
            r"(?:nebenamtliche[rn]?\s+)?Bundesrichter(?:in)?\s+"
            r"(.+?),\s+(?:als\s+)?Einzelrichter(?:in)?"
        )
        
        
        
        
        
        
        flags = 0
    elif language == "fr":
        pattern = (
            r"(?:M\.|Mme)\s+le Juge fédéral\s+(.+?),\s+"
            r"en qualité de juge unique"
        )
        flags = re.IGNORECASE
    elif language == "it":
        pattern = (
            r"Giudice federale\s+(.+?),\s+"
            r"in qualità di giudice unic[oa]"
        )
        flags = re.IGNORECASE
    else:
        return None

    match = re.search(pattern, block, flags)
    return match.group(1).strip() if match else None


def extract_judges(
    block: str | None, language: str | None
) -> tuple[str | None, list[str], bool, str | None]:
    """Extrahiert Richter mit der robusten Bereinigungslogik des Vorbild-Notebooks."""
    if not block or not language:
        return None, [], False, None

    president = extract_president(block, language)

    block_corrections = {
        "Muschietti. Bundesrichter von Felten":
            "Muschietti, Bundesrichter von Felten",
        "Bundesrichterinnen Hohl und Kiss":
            "Bundesrichterinnen Hohl, Kiss",
    }
    for wrong, correct in block_corrections.items():
        block = block.replace(wrong, correct)
    
    # Dominique Hänni eindeutig von Julia Hänni unterscheiden.
    # Dominique ist nebenamtliche Bundesrichterin und wird in der
    # Richter-Masterliste als "D. Hänni" geführt.
    block = re.sub(
        r"nebenamtliche\s+Bundesrichterin\s+(?:D\.\s*)?Hänni(?:\s+D\.)?",
        "nebenamtliche Bundesrichterin D. Hänni",
        block,
        flags=re.IGNORECASE,
    )
    
    president = extract_president(block, language)

    

    

    if language == "de":
        judge_text = re.split(
            r"Gerichtsschreiber(?:in)?", block, maxsplit=1
        )[0]
        judge_text = re.sub(
            r",\s*Präsident(?:in)?\s+Bundesrichter(?:in)?\s+",
            ", ", judge_text, flags=re.IGNORECASE,
        )
        judge_text = re.sub(
            r"Bundesrichterinnen,\s*", "", judge_text, flags=re.IGNORECASE
        )
        judge_text = re.sub(
            r"\s+(?=nebenamtliche[rn]?\s+Bundesrichter)",
            ", ", judge_text, flags=re.IGNORECASE,
        )
        judge_text = re.sub(
            r"(?:nebenamtliche[rn]?\s+)?Bundes?richter(?:in|innen)?\s+",
            "", judge_text, flags=re.IGNORECASE,
        )
        roles = {
            "Präsident", "Präsidentin", "Bundesrichter", "Bundesrichterin",
            "präsidierendes Mitglied", "als präsidierendes Mitglied",
            "präsidierendes Miglied", "als Einzelrichter",
            "als Einzelrichterin", "Einzelrichter", "Einzelrichterin",
            "als Instruktionsrichterin", "als Instruktionsrichter",
            "als präsidierendes Miglied", "präsisierendes Mitglied",
        }

    elif language == "fr":
        judge_text = re.split(
            r"Greffi(?:er|ère)", block, maxsplit=1, flags=re.IGNORECASE
        )[0]
        judge_text = re.sub(
            r"^(?:"
            r"Mmes\s+et\s+M\.\s+la\s+Juge\s+fédérale,?\s*|"
            r"MM\.?\s+et\s+Mme\s+et\s+les\s+Juges\s+fédéraux,?\s*|"
            r"MM\.?\s+et\s+les\s+Juges\s+fédéraux|"
            r"M\.\s+et\s+Mmes?(?:\s+et)?\s+les\s+Juge?s?\s+fédéraux,?\s*|"
            r"M\.\s+et\s+Mmes\s+les\s+Juges\s+fédéraux|"
            r"MM\.?\s+et\s+Mmes\s+les\s+Juges\s+fédéraux|"
            r"Mmes\s+et\s+M\.\s+les\s+Juge?s?\s+fédéra(?:ux|aux),?\s*|"
            r"Mme\s+et\s+(?:M\.\s+)?les\s+Juge?s?\s+fédéraux,?\s*|"
            r"MM\.?\s+et\s+Mme\s+les\s+Juge?s?\s+fédéraux,?\s*|"
            r"MM\.?\s+les\s+Juge?s?\s+fédéraux,?\s*|"
            r"Mmes?\s+et\s+MM\.?\s+les\s+Juge?s?\s+fédéraux,?\s*|"
            r"Mmes?\s+et\s+MM\.\s+les\s+Juges\s+fédéraux|"
            r"Mmes\s+et\s+M\.\s+les\s+Juges\s+fédéraux|"
            r"MM\.\s+et\s+Mme\s+les\s+Juges\s+fédéraux|"
            r"Mmes\s+les\s+Juges\s+fédéral(?:es|aux)|"
            r"MM\.?\s+les\s+Juges\s+fédéraux|"
            r"Mme\s+les?\s+Juges?\s+fédéraux|"
            r"Mme\s+la\s+Juge\s+fédérale|"
            r"M\.\s+le\s+Juge\s+fédéral|"
            r"les\s+Juges\s+fédéraux"
            r")\s+",
            "", judge_text, flags=re.IGNORECASE,
        )
        judge_text = re.sub(
            r",?\s*(?:présidente?|juge\s+présidant|juge\s+suppléante?)\s*,?",
            ", ", judge_text, flags=re.IGNORECASE,
        )
        judge_text = re.sub(r"\s+et\s+", ", ", judge_text)
        roles = {
            "Président", "Présidente", "Juge présidant", "Juge présidante",
            "en qualité de juge unique", "en qualité de juge instructrice",
            "en qualité de Juge unique", "en qualité de Juge instructeur",
            "en qualité de juge instructeur", "Juge unique", "en qualité de",
            "MM", "Juge instructrice", "membre présidant",
        }

    else:
        judge_text = re.split(
            r"Cancellier(?:e|a)", block, maxsplit=1
        )[0]
        judge_text = re.sub(
            r"Giudic(?:e|i)\s+federal(?:e|i)\s+", "", judge_text
        )
        judge_text = re.sub(
            r"\s+e\s+", ", ", judge_text, flags=re.IGNORECASE
        )
        judge_text = re.sub(
            r",?\s*(?:Giudice\s+Presidente|Giudice\s+supplente)\s*,?",
            ", ", judge_text, flags=re.IGNORECASE,
        )
        roles = {
            "Presidente", "Giudice presidente", "Giudice unico",
            "in qualità di giudice unico", "in qualità di giudice unica",
            "Giudice dell'istruzione",
        }

    judges = [
        re.sub(
            r"^(?:"
            r"MM\.?\s+et\s+Mme\s+et\s+les\s+Juges\s+fédéraux|"
            r"Mmes?\s+les\s+Juges\s+fédéraux|"
            r"MM\.?\s+les\s+Juges\s+fédéraux|"
            r"Mme\s+les?\s+Juges?\s+fédéraux|"
            r"Mme\.?\s+les?\s+Juges?\s+fédérales?|"
            r"Mme?\.?\s+la\s+Juge\s+fédérale|"
            r"(?:M\.\s+la\s+Juge|MM\.?\s+et\s+Mme\s+les\s+Juges)\s+fédérales?|"
            r"M\.\s+le\s+Juge(?:\s+fédéral)?|"
            r"nebenamtlicher\s+Bundesrichter|"
            r"Bundesrichter(?:in)?|"
            r"Bundesricher(?:in)?|"
            r"Bundesricherin|"
            r"Bundesrichtger|"
            r"Bundesricherin|"
            r"Bunesrichter|"
            r"Bundesricherin|"
            r"Bundsrichter(?:in)?|"
            r"Juge\s+instruct(?:eur|rice)|"
            r"Mme\.?"
            r")\s+",
            "", part.strip(" ."), flags=re.IGNORECASE,
        )
        for part in judge_text.split(",")
        if part.strip(" .") and part.strip(" .") not in roles
    ]

    judges = [
        re.sub(
            r"^(?:als\s+)?präsidierendes\s+Mitglied\s+|"
            r"\s+als\s+präsidierendes\s+Mitglied$",
            "", name, flags=re.IGNORECASE,
        ).strip()
        for name in judges
    ]

    judges = [
        unicodedata.normalize("NFC", name)
        .replace("\u00a0", " ")
        .replace("\u200b", "")
        .strip(" .;,")
        for name in judges
    ]
    judges = [name for name in judges if name]
    judges = [NAME_CORRECTIONS.get(name, name) for name in judges]
    president = NAME_CORRECTIONS.get(president, president)

    # Doppelte Namen vermeiden, Reihenfolge aber erhalten.
    judges = list(dict.fromkeys(judges))

    if president and president not in judges:
        judges.insert(0, president)
    elif president and judges and judges[0] != president:
        judges.remove(president)
        judges.insert(0, president)

    single_judge_name = extract_single_judge(block, language)
    single_judge_name = NAME_CORRECTIONS.get(single_judge_name, single_judge_name)
    single_judge = len(judges) == 1 or single_judge_name is not None
    if single_judge and single_judge_name is None and judges:
        single_judge_name = judges[0]

    return president, judges, single_judge, single_judge_name

def legacy_legal_area(decision_sign: str) -> str:
    match = re.search(r"^(\d+)[A-Z]_", decision_sign)

    return {
        "1": "Öffentliches Recht",
        "2": "Öffentliches Recht",
        "8": "Öffentliches Recht",
        "9": "Öffentliches Recht",
        "4": "Zivilrecht",
        "5": "Zivilrecht",
        "6": "Strafrecht",
        "7": "Strafrecht",
    }.get(match.group(1), "AndererBereich") if match else "Mistake_Legalara"


def load_external_matchers():
    sys.path.insert(0, str(SCRIPT_DIR))
    try:
        from ext_script_decision_match import search_decisions
        from ext_script_beschwerden_match import search_verfahrensbeteiligte
    except ImportError as exc:
        raise RuntimeError(
            "Die Dateien ext_script_decision_match.py und "
            "ext_script_beschwerden_match.py müssen im gleichen Ordner wie dieses Skript liegen."
        ) from exc
    return search_decisions, search_verfahrensbeteiligte


def analyse_files(files: list[Path], judges_file: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    search_decisions, search_verfahrensbeteiligte = load_external_matchers()
    results = []

    for fallback_id, file in enumerate(sorted(files), start=1):
        soup = BeautifulSoup(file.read_text(encoding="utf-8", errors="replace"), "lxml")
        decision_tag = soup.find("b")
        decision_sign = decision_tag.get_text(strip=True) if decision_tag else ""
        filename_match = re.search(r"urteil__(\d{1,5})__", file.name)
        filename_id = int(filename_match.group(1)) if filename_match else fallback_id
        content = soup.select_one("#highlight_content")
        text = content.get_text("\n", strip=True) if content else ""
        text = re.sub(r"\s+", " ", text).strip()
        for wrong, correct in TEXT_CORRECTIONS.items():
            text = text.replace(wrong, correct)

        language = detect_language(text)
        block = extract_panel_block(text, language)
        president, judges, single_judge, single_judge_name = extract_judges(block, language)
        other_judges = [name for name in judges if name != president]
        parties = search_verfahrensbeteiligte(soup, file.name)
        rss_metadata = RSS_METADATA.get(file.name, {})
        legal_area_code = rss_metadata.get("legal_area_code")
        legal_area_name = rss_metadata.get("legal_area")
        # Bei --analyse-only fehlen RSS-Metadaten; dann bleibt die bisherige
        # grobe Einteilung als Rückfalllösung erhalten.
        if legal_area_name is None:
            legal_area_name = legacy_legal_area(decision_sign)

        results.append({
            "id_filename": filename_id, "filename": file.name,
            "decision_sign": decision_sign, "sprache": language,
            "besetzung_block": block, "praesident": president,
            "einzelrichter": single_judge,
            "einzelrichter_name": single_judge_name,
            "weiterer_richter1": other_judges[0] if len(other_judges) > 0 else None,
            "weiterer_richter2": other_judges[1] if len(other_judges) > 1 else None,
            "weiterer_richter3": other_judges[2] if len(other_judges) > 2 else None,
            "weiterer_richter4": other_judges[3] if len(other_judges) > 3 else None,
            "entscheid": search_decisions(soup),
            "legal_area_code": legal_area_code, "legal area": legal_area_name,
            "beschwerdefuehrer": parties["beschwerdefuehrer_staat"],
            "beschwerdegegner": parties["beschwerdegegner_staat"],
        })

    raw = pd.DataFrame(results)
    if raw.empty:
        return raw, pd.DataFrame()

    long = raw.melt(
        id_vars=["id_filename", "decision_sign", "entscheid", "legal_area_code", "legal area", "beschwerdefuehrer", "beschwerdegegner"],
        value_vars=["praesident", "weiterer_richter1", "weiterer_richter2", "weiterer_richter3", "weiterer_richter4"],
        var_name="funktion", value_name="name",
    )
    judges = pd.read_excel(judges_file)
    merged = long.merge(judges, on="name", how="left").drop(columns="party_raw", errors="ignore")
    
    cleaned = merged[[
        "id_filename", "decision_sign", "name", "both_names", "party", "legal_area_code", "legal area", "entscheid",
        "beschwerdefuehrer", "beschwerdegegner",
    ]].dropna(subset=["name"])
    

     
    
    cleaned = cleaned[cleaned["name"].str.strip() != ""]
    return raw, cleaned


def save_outputs(cleaned: pd.DataFrame, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    cleaned.to_csv(output_dir / "decisions_data_structured.csv", index=False, encoding="utf-8-sig")
    cleaned.to_excel(output_dir / "decisions_date_structured.xlsx", index=False)

    counts = cleaned.groupby("id_filename")["name"].count()
    three_ids = counts[counts == 3].index
    three_judges = cleaned[cleaned["id_filename"].isin(three_ids)].sort_values("id_filename")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Richter"
    sheet.append(three_judges.columns.tolist())
    previous_id = None
    for _, row in three_judges.iterrows():
        current_id = row["id_filename"]
        if previous_id is not None and current_id != previous_id:
            sheet.append([])
        sheet.append(row.tolist())
        previous_id = current_id
    workbook.save(output_dir / "richter_nach_entscheid-3Richter.xlsx")


def save_yearly_json(
    raw: pd.DataFrame,
    cleaned: pd.DataFrame,
    publication_day: str,
    output_dir: Path,
) -> Path:
    """Ergänzt die Jahresdatei um die heute ausgewerteten Entscheide.

    Das Aktenzeichen dient als stabiler Schlüssel. Ein erneuter Lauf für den
    gleichen Tag aktualisiert die betreffenden Einträge, erzeugt aber keine
    Duplikate. Bereits gespeicherte Urteile anderer Tage bleiben erhalten.
    """
    year = publication_day[:4]
    publication_date = (
        f"{publication_day[:4]}-{publication_day[4:6]}-{publication_day[6:8]}"
    )
    json_dir = output_dir / "data"
    json_dir.mkdir(parents=True, exist_ok=True)
    json_file = json_dir / f"decisions-{year}.json"

    if json_file.exists():
        try:
            payload = json.loads(json_file.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError) as exc:
            raise RuntimeError(
                f"Die bestehende JSON-Datei ist nicht lesbar: {json_file}"
            ) from exc
    else:
        payload = {"year": int(year), "updated_at": None, "decisions": {}}

    decisions = payload.setdefault("decisions", {})
    if not isinstance(decisions, dict):
        raise RuntimeError(
            f"In {json_file} muss 'decisions' ein JSON-Objekt sein."
        )

    raw_by_id = raw.set_index("id_filename", drop=False)
    for filename_id, judge_rows in cleaned.groupby("id_filename", sort=False):
        if filename_id not in raw_by_id.index:
            continue
        decision_row = raw_by_id.loc[filename_id]
        # Sicherheit für den theoretischen Fall doppelter IDs.
        if isinstance(decision_row, pd.DataFrame):
            decision_row = decision_row.iloc[0]

        decision_sign = str(decision_row.get("decision_sign", "")).strip()
        if not decision_sign:
            continue

        president = decision_row.get("praesident")
        single_judge = bool(decision_row.get("einzelrichter", False))
        judges = []
        for _, judge in judge_rows.iterrows():
            name = str(judge["name"]).strip()
            both_names = judge.get("both_names")
            if pd.isna(both_names):
                print(f"WARNUNG: Richter nicht in Masterliste gefunden: {name}")
                continue

            full_name = str(both_names).strip()
           
           
           
            if single_judge:
                role = "einzelrichter"
            elif pd.notna(president) and name == str(president).strip():
                role = "praesident"
            else:
                role = "mitglied"

            party = judge.get("party")
            judges.append({
                "name": name,
                "full_name": full_name,
                "party": None if pd.isna(party) else str(party),
                "role": role,
            })

        def json_value(value):
            return None if pd.isna(value) else value

        old = decisions.get(decision_sign, {})
        expected = sum(pd.notna(decision_row.get(field)) for field in
                       ["praesident", *[f"weiterer_richter{i}" for i in range(1, 5)]])
        if old.get("judge_count", 0) > len(judges) or (complete(old) and len(judges) <= old.get("judge_count", 0)):
            continue
        decisions[decision_sign] = {
            "publication_date": old.get("publication_date") or publication_date,
            "richter_vollstaendig": len(judges) == expected and expected > 0,
            "legal_area_code": json_value(decision_row.get("legal_area_code")) or old.get("legal_area_code"),
            "legal_area": (json_value(decision_row.get("legal area")) if pd.notna(decision_row.get("legal_area_code")) else old.get("legal_area") or json_value(decision_row.get("legal area"))),
            "entscheid": json_value(decision_row.get("entscheid")),
            "beschwerdefuehrer": json_value(decision_row.get("beschwerdefuehrer")),
            "beschwerdegegner": json_value(decision_row.get("beschwerdegegner")),
            "judge_count": len(judges),
            "judges": judges,
        }

    payload["year"] = int(year)
    payload["updated_at"] = datetime.now(ZoneInfo("Europe/Zurich")).isoformat(
        timespec="seconds"
    )
    # Sortierte Aktenzeichen ergeben übersichtliche, stabile Git-Änderungen.
    payload["decisions"] = dict(sorted(decisions.items()))

    temporary_file = json_file.with_suffix(".json.tmp")
    temporary_file.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    temporary_file.replace(json_file)
    return json_file


def report_quality(raw: pd.DataFrame, cleaned: pd.DataFrame) -> None:
    if raw.empty:
        return
    all_ids = set(raw["id_filename"])
    recognised_ids = set(cleaned["id_filename"])
    print(f"Ausgewertete Entscheide: {len(raw)}")
    print(f"Gespeicherte Richterzeilen: {len(cleaned)}")
    print(f"Entscheide ohne erkannte Richter: {sorted(all_ids - recognised_ids)}")
    if "party" in cleaned:
        missing = sorted(cleaned.loc[cleaned["party"].isna(), "name"].unique())
        print(f"Richternamen ohne Parteizuordnung: {missing}")
    counts = cleaned.groupby("id_filename")["name"].count().value_counts().sort_index()
    print("Anzahl Entscheide nach Gremiumsgrösse:")
    print(counts.to_string())


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--date", help="Datum im Format JJJJMMTT; Standard: letzte fünf Kalendertage in der Schweiz")
    parser.add_argument("--html-dir", type=Path, default=DEFAULT_HTML_DIR)
    parser.add_argument("--output-dir", type=Path, default=SCRIPT_DIR)
    parser.add_argument("--judges-file", type=Path, default=DEFAULT_JUDGES_FILE)
    parser.add_argument("--analyse-only", action="store_true", help="Bereits gespeicherte HTML-Dateien nur auswerten")
    return parser.parse_args()


def enrich_existing(day, output_dir, metadata):
    path = output_dir / "data" / f"decisions-{day[:4]}.json"
    if not path.exists():
        return
    payload = json.loads(path.read_text(encoding="utf-8"))
    changed = False
    for key, entry in payload["decisions"].items():
        info = metadata.get(case_key(key))
        if info and any(entry.get(k) != v for k, v in info.items()):
            entry.update(info)
            changed = True
    if changed:
        payload["updated_at"] = datetime.now(ZoneInfo("Europe/Zurich")).isoformat(timespec="seconds")
        temp = path.with_suffix(".json.tmp")
        temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        temp.replace(path)


def main() -> int:
    args = parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    if args.date:
        start = datetime.strptime(args.date, "%Y%m%d")
        days = [start.strftime("%Y%m%d")]
    else:
        start = datetime.now(ZoneInfo("Europe/Zurich"))
        days = [(start - timedelta(days=i)).strftime("%Y%m%d") for i in range(4, -1, -1)]
    if args.analyse_only and not args.date:
        raise SystemExit("--analyse-only erfordert --date")
    if not args.judges_file.exists():
        raise FileNotFoundError(args.judges_file)
    log_file = args.output_dir / "scraping_log.txt"
    # Zuerst die Newsseiten verarbeiten, unabhängig von der RSS-Verfügbarkeit.
    failed = False
    stopped = False
    for day in days:
        try:
            existing = existing_decisions(day, args.output_dir)
            RSS_METADATA.clear()
            if args.analyse_only:
                stamp = f"_{day[:4]}_{day[4:6]}_{day[6:]}__"
                files = sorted(f for f in args.html_dir.glob("*.html") if stamp in f.name)
            else:
                files, stopped, errors = scrape_daily_decisions(day, args.html_dir, log_file, existing, {})
                failed = failed or bool(errors)
            if files:
                raw, cleaned = analyse_files(files, args.judges_file)
                save_outputs(cleaned, args.output_dir)
                save_yearly_json(raw, cleaned, day, args.output_dir)
                report_quality(raw, cleaned)
        except AccessStopped as exc:
            log(f"Abruf eingestellt am {day}: {exc}", log_file)
            failed = True
            stopped = True
        except Exception as exc:
            log(f"Fehler am {day}: {exc}", log_file)
            failed = True
        if stopped:
            break
    try:
        metadata = {} if stopped else rss_metadata()
        for day in days:
            enrich_existing(day, args.output_dir, metadata)
    except Exception as exc:
        log(f"RSS-Rechtsgebiete derzeit nicht verfügbar: {exc}", log_file)
    # Erfolgreiche Teilergebnisse sollen auch bei einzelnen Tagesfehlern gepusht werden.
    if failed:
        log("WARNUNG: Mindestens ein Tag konnte nicht vollständig verarbeitet werden.", log_file)
        print("::warning::Urteilsabruf unvollständig; Teilergebnisse werden gespeichert.", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
