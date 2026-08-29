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
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin
from zoneinfo import ZoneInfo

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook
from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


BASE_URL = "https://search.bger.ch/ext/eurospider/live/de/php/aza/http/index_aza.php"
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_HTML_DIR = SCRIPT_DIR / "aktuelle_urteile"
DEFAULT_JUDGES_FILE = SCRIPT_DIR / "list_federal_judges_party.xlsx"


def log(message: str, log_file: Path) -> None:
    print(message)
    with log_file.open("a", encoding="utf-8") as handle:
        handle.write(message + "\n")


def current_swiss_date() -> str:
    """Aktuelles Datum in der für die Bundesgerichts-URL nötigen Form."""
    return datetime.now(ZoneInfo("Europe/Zurich")).strftime("%Y%m%d")


def scrape_daily_decisions(
    day: str,
    output_dir: Path,
    log_file: Path,
    *,
    headless: bool = True,
) -> list[Path]:
    """Lädt alle auf der Newsseite eines Tages verlinkten Entscheide herunter."""
    output_dir.mkdir(parents=True, exist_ok=True)
    news_url = f"{BASE_URL}?date={day}&lang=de&mode=news"
    saved_files: list[Path] = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=headless)
        page = browser.new_page()
        try:
            page.goto(news_url, wait_until="networkidle", timeout=60_000)
            locator = page.locator('a[href*="highlight_docid="]')
            links: list[str] = []
            for index in range(locator.count()):
                href = locator.nth(index).get_attribute("href")
                if href:
                    absolute = urljoin(page.url, href)
                    if absolute not in links:
                        links.append(absolute)

            if not links:
                log(f"Am {day} wurden keine Urteile gefunden.", log_file)
                return []

            year, month, day_number = day[:4], day[4:6], day[6:8]
            for number, decision_url in enumerate(links, start=1):
                match = re.search(r"\d+[A-Z]_[0-9]+-\d{4}", decision_url)
                decision_sign = match.group(0) if match else f"unbekannt_{number}"
                safe_sign = re.sub(r"[^0-9A-Za-z_-]+", "_", decision_sign)
                filename = (
                    f"urteil__{number}__{number}_{year}_{month}_{day_number}__"
                    f"{safe_sign}.html"
                )
                target = output_dir / filename

                try:
                    page.goto(decision_url, wait_until="networkidle", timeout=60_000)
                    target.write_text(page.content(), encoding="utf-8")
                    saved_files.append(target)
                    log(f"Gespeichert: {target.name}", log_file)
                except PlaywrightTimeoutError:
                    log(f"Zeitüberschreitung beim Urteil: {decision_url}", log_file)
        finally:
            browser.close()

    return saved_files


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
            r"Bundesrichter(?:in)?\s+(.+?),\s*"
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
        pattern = r"Bundesrichter(?:in)?\s+(.+?),\s+als Einzelrichter(?:in)?"
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
            "als Einzelrichterin", "als Instruktionsrichterin",
            "als Instruktionsrichter", "als präsidierendes Miglied",
            "Einzelrichterin", "präsisierendes Mitglied",
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

def legal_area(decision_sign: str) -> str:
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
            "entscheid": search_decisions(soup), "legal area": legal_area(decision_sign),
            "beschwerdefuehrer": parties["beschwerdefuehrer_staat"],
            "beschwerdegegner": parties["beschwerdegegner_staat"],
        })

    raw = pd.DataFrame(results)
    if raw.empty:
        return raw, pd.DataFrame()

    long = raw.melt(
        id_vars=["id_filename", "decision_sign", "entscheid", "legal area", "beschwerdefuehrer", "beschwerdegegner"],
        value_vars=["praesident", "weiterer_richter1", "weiterer_richter2", "weiterer_richter3", "weiterer_richter4"],
        var_name="funktion", value_name="name",
    )
    judges = pd.read_excel(judges_file)
    merged = long.merge(judges, on="name", how="left").drop(columns="party_raw", errors="ignore")
    cleaned = merged[[
        "id_filename", "decision_sign", "name", "party", "legal area", "entscheid",
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
            if single_judge:
                role = "einzelrichter"
            elif pd.notna(president) and name == str(president).strip():
                role = "praesident"
            else:
                role = "mitglied"

            party = judge.get("party")
            judges.append({
                "name": name,
                "party": None if pd.isna(party) else str(party),
                "role": role,
            })

        def json_value(value):
            return None if pd.isna(value) else value

        decisions[decision_sign] = {
            "publication_date": publication_date,
            "legal_area": json_value(decision_row.get("legal area")),
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
    parser.add_argument("--date", help="Datum im Format JJJJMMTT; Standard: heutiges Datum in der Schweiz")
    parser.add_argument("--html-dir", type=Path, default=DEFAULT_HTML_DIR)
    parser.add_argument("--output-dir", type=Path, default=SCRIPT_DIR)
    parser.add_argument("--judges-file", type=Path, default=DEFAULT_JUDGES_FILE)
    parser.add_argument("--show-browser", action="store_true", help="Chromium während des Scrapens anzeigen")
    parser.add_argument("--analyse-only", action="store_true", help="Bereits gespeicherte HTML-Dateien nur auswerten")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    day = args.date or current_swiss_date()
    if not re.fullmatch(r"\d{8}", day):
        raise SystemExit("--date muss das Format JJJJMMTT haben, zum Beispiel 20260826.")
    log_file = args.output_dir / "scraping_log.txt"

    if args.analyse_only:
        files = sorted(args.html_dir.glob("*.html"))
    else:
        files = scrape_daily_decisions(day, args.html_dir, log_file, headless=not args.show_browser)

    if not files:
        print("Keine HTML-Dateien zum Auswerten vorhanden; es wurden keine Tabellen erzeugt.")
        return 0
    if not args.judges_file.exists():
        raise FileNotFoundError(f"Richterdatei nicht gefunden: {args.judges_file}")

    raw, cleaned = analyse_files(files, args.judges_file)
    save_outputs(cleaned, args.output_dir)
    json_file = save_yearly_json(raw, cleaned, day, args.output_dir)
    report_quality(raw, cleaned)
    print(f"CSV- und Excel-Dateien gespeichert in: {args.output_dir}")
    print(f"Jahres-JSON gespeichert in: {json_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
