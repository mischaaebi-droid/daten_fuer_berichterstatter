"""
Extraktion und Klassifikation der Verfahrensbeteiligten in Strafrechtsurteilen.

Produktive Version der im Notebook test_verfahrensbeteiligte.ipynb
entwickelten und kontrollierten Logik.

Verwendung:
    from ext_script_verfahrensbeteiligte import search_verfahrensbeteiligte

    beteiligte = search_verfahrensbeteiligte(soup, file.name)
"""

import re


staat_muster = {

    "staatsanwaltschaft": [
        r"\bstaatsanwaltschaft\b",
        r"\boberstaatsanwaltschaft\b",
        r"\bgeneralstaatsanwaltschaft\b",
        r"\bbundesanwaltschaft\b",
        r"\bjugendanwaltschaft\b",
        r"\boberjugendanwaltschaft\b",
        r"\bjugendstaatsanwaltschaft\b",
        r"\bstaatsanwältin\b",
        r"\bstaatsanwalt\b",
        r"\bprocureur(?:e)?\b",
    
        # fehlerhafte Schreibweisen in Originaltexten
        r"\bobersta\s*atsanwaltschaft\b",
        r"\bgener\s*alstaatsanwaltschaft\b",
        r"\bstaatsanwaltscha\s*ft\b",
        r"\boberstaatsanw\s*altschaft\b",
        r"\boberstaatsanwaltschaf\s*t\b",
    
        r"\bstadtrichteramt\b",
        r"\bstatthalteramt\b",
        r"\buntersuchungsamt\b",
    
        r"\bministère public\b",
        r"\bparquet général\b",
        r"\bprocureur général\b",
        r"\bprocureure générale\b",
        r"\bprocureur général adjoint\b",
    
        r"\bministero pubblico\b",
        r"\bprocura pubblica\b",
    ],

    "gericht_richter": [
        r"\bkantonsgerichts?\b",
        r"\bobergerichts?\b",
        r"\bbezirksgerichts?\b",
        r"\bregionalgerichts?\b",
        r"\bstrafgerichts?\b",
        r"\bbundesstrafgerichts?\b",
        r"\bappellationsgerichts?\b",
        r"\bverwaltungsgerichts?\b",
        r"\bzwangsmassnahmengerichts?\b",
        r"\banklagekammer\b",
    
        r"\bkreisrichter(?:in)?\b",
        r"\bpolizeirichter(?:in)?\b",
        r"\bamtsgerichtspräsident(?:in)?\b",
        r"\bgerichtspräsident(?:in)?\b",
        r"\boberrichter(?:in)?\b",
        r"\brichter(?:in)?\b",
        r"\brichteramt\b",
        r"\bgerichtsschreiber(?:in)?\b",
        r"\bkantonsgerichtspräsidium\b",
        r"\bobergerichtspräsidium\b",
    
        r"\btribunal\b",
        r"\bcour de justice\b",
        r"\bchambre pénale\b",
        r"\bjuge\b",
        r"\bobergerichtspräsident(?:in)?\b",
        r"\bkantonsgerichtspräsident(?:in)?\b",
        r"\bbezirksgerichtspräsident(?:in)?\b",
        r"\bstrafgerichtspräsident(?:in)?\b",


        
    ],

    "polizei": [
        r"\bkantonspolizei\b",
        r"\bstadtpolizei\b",
        r"\bpolizei\b",
        r"\bpolice\b",
    ],

    "justizvollzug": [
        r"\bjustizvollzug\b",
        r"\bbewährungsdienst\b",
        r"\bbewährungs- und vollzugsdienst\b",
        r"\bvollzugs- und bewährungsdienst\b",
        r"\bbewährungs- und vollzugsdienste\b",
        r"\bstraf- und massnahmenvollzug\b",
        r"\bstrafvollzug\b",
        r"\bmassnahmenvollzug\b",
        r"\bétablissement pénitentiaire\b",
    ],

    "steuer_zoll": [
        r"\beidgenössische steuerverwaltung\b",
        r"\badministration fédérale des contributions\b",
        r"\bamministrazione federale delle contribuzioni\b",
        r"\bbundesamt für zoll und grenzsicherheit\b",
        r"\boffice fédéral de la douane\b",
    ],

    "migration": [
        r"\bmigrationsamt\b",
        r"\bamt für migration\b",
        r"\bstaatssekretariat für migration\b",
    ],

    "andere_behoerde": [
        r"\bregierungsrat\b",
        r"\bconseil d'état\b",
        r"\bconsiglio di stato\b",
    
        r"\bsicherheitsdirektion\b",
        r"\bjustiz- und sicherheitsdepartement\b",
        r"\bsicherheits- und justizdepartement\b",
        r"\bdepartement des innern\b",
        r"\bdepartement\b",
        r"\bdirektion der justiz und des innern\b",
        r"\bdipartimento delle istituzioni\b",
    
        r"\bgemeinde\b",
        r"\beinwohnergemeinde\b",
        r"\bcommune\b",
        r"\bstadt\b",
        r"\bsozialbehörde\b",
    
        r"\bkindes- und erwachsenenschutzbehörde\b",
        r"\bkesb\b",
    
        r"\bservice des contraventions\b",
        r"\bsezione della circolazione\b",
        r"\beidgenössische spielbankenkommission\b",
        r"\bbundesamt für umwelt\b",
        r"\bchef d['’]office\b",
    ],
}


def beschwerdefuehrer_text_bereinigen(text):

    # c/o-Adressen entfernen
    text = re.sub(
        r"\bc/o\s+[^,]+,",
        "",
        text,
        flags=re.IGNORECASE
    )

    # französische Aufenthaltsangaben in Gefängnissen entfernen
    text = re.sub(
        r"\bactuellement détenu(?:e)?\s+à\s+[^,]+,",
        "",
        text,
        flags=re.IGNORECASE
    )


    # KESB als Vertretung einer privaten Person entfernen
    text = re.sub(
        r"\bhandelnd durch\b.*?\bKESB\b[^,]*,",
        "",
        text,
        flags=re.IGNORECASE
    )
    



    

    return text










def staatliche_stelle_erkennen(text):

    for staat_typ, muster_liste in staat_muster.items():

        for muster in muster_liste:

            if re.search(
                muster,
                text,
                flags=re.IGNORECASE
            ):
                return "staat", staat_typ

    return "privat", None







    

# ============================================================
# Sprache erkennen
# ============================================================

def sprache_erkennen(text):

    if (
        "Verfahrensbeteiligte" in text
        or re.search(
            r"\bBeschwerdeführer(?:in)?\b",
            text,
            flags=re.IGNORECASE
        )
    ):
        return "de"

    

    elif (
        re.search(
            r"\bParticipants?\s+à\s+la\s+procédure\b",
            text,
            flags=re.IGNORECASE
        )
        or re.search(
            r"\brecourant(?:e|s|es)?\b",
            text,
            flags=re.IGNORECASE
        )
    ):
        return "fr"
    
    


    elif "Partecipanti al procedimento" in text:
        return "it"

    else:
        return None


# ============================================================
# Beteiligtenblock extrahieren
# ============================================================

def beteiligtenblock_finden(text, sprache):

    if sprache == "de":

        treffer = re.search(
            r"Verfahrensbeteiligte\s+(.*?)\s+Gegenstand",
            text,
            flags=re.IGNORECASE | re.DOTALL
        )

    elif sprache == "fr":

        treffer = re.search(
            r"Participants?\s+à\s+la\s+procédure\s+(.*?)\s+Objet",
            text,
            flags=re.IGNORECASE | re.DOTALL
        )

    elif sprache == "it":

        treffer = re.search(
            r"Partecipanti al procedimento\s+(.*?)\s+Oggetto",
            text,
            flags=re.IGNORECASE | re.DOTALL
        )

    else:
        return None

    if treffer:
        return treffer.group(1).strip()

    return None


# ============================================================
# Parteiseiten separat extrahieren
# ============================================================

def parteiseiten_trennen(block, sprache):

    # --------------------------------------------------------
    # 1. Trennwort bestimmen
    # --------------------------------------------------------

    if sprache == "de":
        trennwort = r"\bgegen\b"

    elif sprache == "fr":
        trennwort = r"\bcontre\b"

    elif sprache == "it":
        trennwort = r"\bcontro\b"

    else:
        return None, None





    # --------------------------------------------------------
    # 3. Mehrere gegen / contre / contro = verbundenes Verfahren
    # --------------------------------------------------------

    if len(
        re.findall(
            trennwort,
            block,
            flags=re.IGNORECASE
        )
    ) > 1:
        return None, None


    # --------------------------------------------------------
    # 4. Block beim gegen / contre / contro teilen
    # --------------------------------------------------------

    teile = re.split(
        trennwort,
        block,
        maxsplit=1,
        flags=re.IGNORECASE
    )

    if len(teile) != 2:
        return None, None


    beschwerdefuehrer_text = teile[0].strip()
    beschwerdegegner_text = teile[1].strip()


    # --------------------------------------------------------
    # 5. Ende des eigentlichen Beschwerdegegners bestimmen
    # --------------------------------------------------------

    if sprache == "de":
        rollenende = r"\bBeschwerdegegner(?:in|innen)?\b"

    elif sprache == "fr":
        rollenende = r"\bintimé(?:e|s|es)?\b"

    elif sprache == "it":
        rollenende = r"\bopponent(?:e|i)?\b"

    else:
        rollenende = None


    # --------------------------------------------------------
    # 6. Alles NACH dem Beschwerdegegner abschneiden
    # --------------------------------------------------------

    if rollenende:

        treffer_liste = list(
            re.finditer(
                rollenende,
                beschwerdegegner_text,
                flags=re.IGNORECASE
            )
        )
    
        if treffer_liste:
            letzter_treffer = treffer_liste[-1]
    
            beschwerdegegner_text = (
                beschwerdegegner_text[:letzter_treffer.end()]
            )


    # --------------------------------------------------------
    # 7. Beide Seiten zurückgeben
    # --------------------------------------------------------

    return beschwerdefuehrer_text, beschwerdegegner_text

def search_verfahrensbeteiligte(soup, filename=None):
    """
    Klassifiziert die Beschwerdeführer- und Beschwerdegegnerseite
    eines Strafrechtsurteils.

    Nur B-Entscheide werden ausgewertet. Für andere Rechtsgebiete
    wird der Status 'NICHT ANWENDBAR' zurückgegeben.

    Rückgabe: Dictionary mit
      - beschwerdefuehrer_staat
      - beschwerdefuehrer_staat_typ
      - beschwerdegegner_staat
      - beschwerdegegner_staat_typ
      - verfahrensbeteiligte_block
      - verfahrensbeteiligte_status
    """

    leer = {
        "beschwerdefuehrer_staat": None,
        "beschwerdefuehrer_staat_typ": None,
        "beschwerdegegner_staat": None,
        "beschwerdegegner_staat_typ": None,
        "verfahrensbeteiligte_block": None,
        "verfahrensbeteiligte_status": None,
    }

    # Nur Strafrechtsurteile (B-Entscheide)
    # Primär Dateiname prüfen; falls keiner übergeben wurde,
    # Aktenzeichen im Urteilstext verwenden.
    if filename is not None:
        ist_strafrecht = bool(
            re.search(r"__\d+B_", filename, flags=re.IGNORECASE)
        )
    else:
        gesamttext = soup.get_text(" ", strip=True)
        ist_strafrecht = bool(
            re.search(r"\b\d+B_\d+/\d{4}\b", gesamttext)
        )

    if not ist_strafrecht:
        leer["verfahrensbeteiligte_status"] = "NICHT ANWENDBAR"
        return leer

    # Gleiche Textgewinnung wie im Testnotebook
    urteilsbereich = soup.select_one("#highlight_content")

    if urteilsbereich:
        text = urteilsbereich.get_text(" ", strip=True)
    else:
        text = soup.get_text(" ", strip=True)

    text = re.sub(r"\s+", " ", text).strip()

    sprache = sprache_erkennen(text)

    if sprache is None:
        leer["verfahrensbeteiligte_status"] = "SPRACHE NICHT ERKANNT"
        return leer

    block = beteiligtenblock_finden(text, sprache)

    if block is None:
        leer["verfahrensbeteiligte_status"] = "BLOCK NICHT ERKANNT"
        return leer

    bf_text, bg_text = parteiseiten_trennen(block, sprache)

    # Komplexe Mehrfachverfahren oder Fälle ohne Gegenpartei
    if bf_text is None or bg_text is None:
        return {
            "beschwerdefuehrer_staat": "NICHT KLASSIFIZIERT",
            "beschwerdefuehrer_staat_typ": None,
            "beschwerdegegner_staat": "NICHT KLASSIFIZIERT",
            "beschwerdegegner_staat_typ": None,
            "verfahrensbeteiligte_block": block,
            "verfahrensbeteiligte_status": "NICHT KLASSIFIZIERT",
        }

    # Sonderfälle auf Beschwerdeführerseite bereinigen
    bf_text_bereinigt = beschwerdefuehrer_text_bereinigen(bf_text)

    bf_staat, bf_staat_typ = staatliche_stelle_erkennen(
        bf_text_bereinigt
    )
    bg_staat, bg_staat_typ = staatliche_stelle_erkennen(
        bg_text
    )

    return {
        "beschwerdefuehrer_staat": bf_staat,
        "beschwerdefuehrer_staat_typ": bf_staat_typ,
        "beschwerdegegner_staat": bg_staat,
        "beschwerdegegner_staat_typ": bg_staat_typ,
        "verfahrensbeteiligte_block": block,
        "verfahrensbeteiligte_status": "KLASSIFIZIERT",
    }