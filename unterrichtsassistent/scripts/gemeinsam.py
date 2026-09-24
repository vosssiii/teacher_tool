# -*- coding: utf-8 -*-
"""
Gemeinsame Bausteine der Skripte: Metadatenkopf, Dateinamen, Pruefwerte.

Kein eigenstaendiges Skript. Wird von den anderen Skripten importiert; das
klappt, weil alle Skripte flach in einem Ordner liegen - im Plugin unter
scripts/, beim Lehrer unter _system/skripte/.
"""

import hashlib
import json
import re
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SCHEMA = 1

# Erlaubte Werte (Kap. 5.3)
FAECHER = ("englisch", "arbeit-recht")
TYPEN = ("arbeitsblatt", "praesentation", "klassenarbeit", "lehrplan",
         "gesetzestext", "fachtext", "vokabelliste", "sonstiges")
HERKUNFT = ("eigen", "fremd", "claude")
STATUS = ("entwurf", "pruefen", "bereit")

# Feste Reihenfolge der Felder im Kopf. Unbekannte Felder werden hinten
# angehaengt statt verworfen - die Lehrkraft darf eigene ergaenzen.
FELDREIHENFOLGE = [
    "schema", "fach", "klasse", "thema", "typ", "beschreibung", "herkunft",
    "quelle", "original", "url", "status", "grafiken", "rechtsstand",
    "gruppe", "dauer", "hilfsmittel", "erstellt", "aktualisiert",
]

# Marker, die eine Datei in den Status "pruefen" zwingen (Kap. 5.4, 6.6)
PRUEFMARKER = re.compile(r"\[(Prüfen|Pruefen|unsicher)\s*:", re.IGNORECASE)

# Platzhalter, die zu_markdown.py fuer Claude hinterlaesst. Steht einer davon
# noch im Text, hat niemand hingesehen - dann darf nichts abgelegt werden.
SKRIPT_PLATZHALTER = [
    (re.compile(r"!\[Bild: beschreiben( oder streichen)?\]\(([^)]+)\)"),
     "Bild noch nicht angesehen: %s – beschreiben oder die Zeile löschen"),
    (re.compile(r"\[Prüfen: Was zeigt das Bild\?"),
     "Bildbeschreibung fehlt noch"),
    (re.compile(r"\[Prüfen: Allen lesbaren Text"),
     "Text im Bild noch nicht abgeschrieben (oder Abschnitt löschen, falls keiner da ist)"),
    (re.compile(r"\[Prüfen: Seite (\d+) ist vermutlich ein Scan"),
     "Scan-Seite %s noch nicht abgeschrieben"),
]


def platzhalter_finden(rumpf):
    """Liste verstaendlicher Meldungen zu noch offenen Skript-Platzhaltern."""
    meldungen = []
    for muster, meldung in SKRIPT_PLATZHALTER:
        for treffer in muster.finditer(rumpf):
            wert = treffer.groups()[-1] if treffer.groups() else None
            meldungen.append(meldung % wert if "%s" in meldung else meldung)
    return meldungen


def heute():
    return date.today().isoformat()


def jetzt():
    return datetime.now().astimezone().isoformat(timespec="seconds")


def pruefwert(pfad):
    """Inhalts-Hash (Kap. 5.6). Bewusst nicht das Dateidatum."""
    h = hashlib.sha256()
    with open(str(pfad), "rb") as datei:
        for block in iter(lambda: datei.read(65536), b""):
            h.update(block)
    return h.hexdigest()


def kurzname(text, maxlaenge=60):
    """Dateiname aus freiem Text: klein, ohne Umlaute, Bindestriche.
    Umlaute werden ausgeschrieben, nicht verschluckt: Kündigung -> kuendigung."""
    text = str(text or "").strip().lower()
    for alt, neu in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss")):
        text = text.replace(alt, neu)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^a-z0-9]+", "-", text).strip("-")
    if len(text) > maxlaenge:
        text = text[:maxlaenge].rsplit("-", 1)[0] or text[:maxlaenge]
    return text or "ohne-titel"


def freier_pfad(pfad):
    """Gibt pfad zurueck, oder - falls belegt - pfad mit -2, -3 ...
    Vorhandenes wird nie ueberschrieben (Designprinzip 6)."""
    pfad = Path(pfad)
    if not pfad.exists():
        return pfad
    nummer = 2
    while True:
        kandidat = pfad.with_name("%s-%d%s" % (pfad.stem, nummer, pfad.suffix))
        if not kandidat.exists():
            return kandidat
        nummer += 1


# --------------------------------------------------------------------------
# Metadatenkopf
# --------------------------------------------------------------------------

def kopf_trennen(text):
    """Teilt eine .md in (Kopf als dict, Rumpf). Ohne Kopf: ({}, text)."""
    if not text.startswith("---"):
        return {}, text
    zeilen = text.split("\n")
    if zeilen[0].strip() != "---":
        return {}, text
    for nummer in range(1, len(zeilen)):
        if zeilen[nummer].strip() == "---":
            roh = "\n".join(zeilen[1:nummer])
            rumpf = "\n".join(zeilen[nummer + 1:]).lstrip("\n")
            try:
                import yaml
                felder = yaml.safe_load(roh) or {}
            except Exception:
                felder = _kopf_notbehelf(roh)
            if not isinstance(felder, dict):
                felder = {}
            # YAML macht aus 2026-09-15 ein Datum - zurueck in Text.
            for schluessel, wert in list(felder.items()):
                if isinstance(wert, (date, datetime)):
                    felder[schluessel] = wert.isoformat()
            return felder, rumpf
    return {}, text


def _kopf_notbehelf(roh):
    """Einfaches Zeilenparsen, falls PyYAML fehlt oder der Kopf kaputt ist."""
    felder = {}
    for zeile in roh.splitlines():
        if ":" in zeile and not zeile.startswith((" ", "\t", "#")):
            schluessel, wert = zeile.split(":", 1)
            wert = wert.strip().strip('"').strip("'")
            felder[schluessel.strip()] = wert or None
    return felder


def _wert_schreiben(wert):
    if wert is None or wert == "":
        return ""
    if isinstance(wert, bool):
        return "ja" if wert else "nein"
    if isinstance(wert, (int, float)):
        return str(wert)
    text = str(wert).replace("\n", " ").strip()
    # Anfuehrungszeichen nur, wo YAML den Wert sonst missversteht. Daten wie
    # 2026-09-24 bleiben ohne - so steht es im Konzept, und so liest es sich
    # fuer die Lehrkraft besser. "ja"/"nein" sind in YAML keine Wahrheitswerte.
    if (re.search(r"[:#\[\]{}&*!|>'\"%@`,]", text) or text[:1] in "-?"
            or text.lower() in ("yes", "no", "true", "false", "null", "on", "off", "y", "n", "~")
            or re.fullmatch(r"\d+([.,]\d+)?", text)):
        return '"%s"' % text.replace("\\", "\\\\").replace('"', '\\"')
    return text


def kopf_schreiben(felder, rumpf):
    """Setzt Kopf und Rumpf zusammen, Felder in fester Reihenfolge."""
    zeilen = ["---"]
    bekannte = [f for f in FELDREIHENFOLGE if f in felder]
    eigene = [f for f in felder if f not in FELDREIHENFOLGE]
    for schluessel in bekannte + eigene:
        wert = _wert_schreiben(felder[schluessel])
        zeilen.append(("%s: %s" % (schluessel, wert)).rstrip())
    zeilen.append("---")
    return "\n".join(zeilen) + "\n\n" + rumpf.lstrip("\n")


def md_lesen(pfad):
    return kopf_trennen(Path(pfad).read_text(encoding="utf-8"))


def md_schreiben(pfad, felder, rumpf):
    Path(pfad).parent.mkdir(parents=True, exist_ok=True)
    Path(pfad).write_text(kopf_schreiben(felder, rumpf), encoding="utf-8")


# --------------------------------------------------------------------------
# Pruefwerte in _system/stand.json
# --------------------------------------------------------------------------

def stand_lesen(wurzel):
    pfad = Path(wurzel) / "_system" / "stand.json"
    if pfad.exists():
        try:
            daten = json.loads(pfad.read_text(encoding="utf-8"))
            daten.setdefault("dateien", {})
            daten.setdefault("aufgenommen", {})
            return daten
        except Exception:
            pass
    return {"schema": SCHEMA, "dateien": {}, "aufgenommen": {}}


def stand_schreiben(wurzel, daten):
    pfad = Path(wurzel) / "_system" / "stand.json"
    pfad.parent.mkdir(parents=True, exist_ok=True)
    pfad.write_text(json.dumps(daten, indent=2, ensure_ascii=False), encoding="utf-8")


def relativ(pfad, wurzel):
    """Pfad relativ zum Arbeitsordner, immer mit / - auch unter Windows."""
    return Path(pfad).resolve().relative_to(Path(wurzel).resolve()).as_posix()


def ausgeben(daten, als_json, textfunktion):
    if als_json:
        print(json.dumps(daten, indent=2, ensure_ascii=False, default=str))
    else:
        print(textfunktion(daten))
