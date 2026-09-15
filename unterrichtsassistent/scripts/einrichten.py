#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Richtet den Arbeitsordner der Lehrkraft ein (KONZEPT.md Kap. 6.1).

Macht alles, was *immer* gleich passieren muss (Designprinzip 3):
Ordner anlegen, Skripte und Vorlagen bereitstellen, Index und Prueflisten
anlegen, Bereitschaft je Befehl feststellen. Was ein Gespraech braucht -
das Interview fuer schule.md, das Umwandeln mitgebrachter Dokumente - macht
der Skill 'einrichten'.

Zwei Regeln, die dieses Skript nie verletzt (Designprinzip 1 und 6):
  - Alles unter _system/ gehoert dem Plugin und wird ueberschrieben.
  - Alles andere wird nur angelegt, wenn es fehlt. Nie ueberschrieben.

Laeuft in beiden Cowork-Umgebungen (Kap. 2.2) und nimmt keine festen Pfade
an - der Arbeitsordner wird uebergeben.

Aufruf:
    python einrichten.py --ordner "/pfad/zu/Unterricht"
    python einrichten.py --ordner . --plugin /pfad/zum/plugin --json
"""

import argparse
import hashlib
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SCHEMA = 1

# Ordner, die zum Arbeitsordner gehoeren (Kap. 4.3).
ORDNER = [
    "_kontext",
    "_system",
    "_system/skripte",
    "_system/vorlagen",
    "eingang",
    "originale",
    "wissensbasis",
    "wissensbasis/englisch",
    "wissensbasis/englisch/bilder",
    "wissensbasis/arbeit-recht",
    "wissensbasis/arbeit-recht/bilder",
    "entwuerfe",
    "ausgabe",
]

# Was jeder Befehl mindestens braucht (Kap. 6.1).
BEDARF = {
    "aufnehmen": [],
    "material": ["_kontext/schule.md", "vorlage"],
    "klassenarbeit": ["_kontext/schule.md", "vorlage",
                      "_kontext/notenschluessel.md",
                      "_kontext/anforderungsbereiche.md",
                      "_kontext/operatoren.md"],
}

KLARTEXT = {
    "_kontext/schule.md": "Angaben zu deiner Schule",
    "_kontext/notenschluessel.md": "dein Notenschlüssel",
    "_kontext/anforderungsbereiche.md": "die Vorgabe zur Punkteverteilung",
    "_kontext/operatoren.md": "deine Operatorenliste",
    "vorlage": "eine Word-Vorlage",
}

INDEX_KOPF = """# Inhaltsverzeichnis der Wissensbasis

Diese Datei pflegt das Plugin. Bitte nicht von Hand bearbeiten – Änderungen
gehen beim nächsten Durchlauf verloren. Was hier steht, stammt aus dem
Metadaten-Kopf der einzelnen Materialien.

| Datei | Fach | Klasse | Typ | Thema | Status | Herkunft | Beschreibung |
| --- | --- | --- | --- | --- | --- | --- | --- |
"""


def jetzt():
    return datetime.now().astimezone().isoformat(timespec="seconds")


def pruefwert(pfad):
    """Inhalts-Hash. Grundlage der Aenderungserkennung (Kap. 5.6) - bewusst
    nicht das Dateidatum, das bei Sync und Kopieren unzuverlaessig ist."""
    h = hashlib.sha256()
    with open(str(pfad), "rb") as datei:
        for block in iter(lambda: datei.read(65536), b""):
            h.update(block)
    return h.hexdigest()


# --------------------------------------------------------------------------
# Schritt 1: Ordner
# --------------------------------------------------------------------------

def ordner_anlegen(wurzel):
    angelegt, vorhanden = [], []
    for name in ORDNER:
        ziel = wurzel / name
        if ziel.exists():
            vorhanden.append(name)
        else:
            ziel.mkdir(parents=True, exist_ok=True)
            angelegt.append(name)
    return {"angelegt": angelegt, "vorhanden": vorhanden}


# --------------------------------------------------------------------------
# Schritt 2: Skripte und Vorlagen bereitstellen
# --------------------------------------------------------------------------

def bereitstellen(plugin_wurzel, wurzel):
    """Kopiert Skripte und Vorlagen in den Arbeitsordner.

    Noetig, weil ${CLAUDE_PLUGIN_ROOT} in Cowork leer ist und das Plugin in
    der lokalen Geraete-Sandbox gar nicht vorliegt (Befund T4, Kap. 4.2).
    Ohne diesen Schritt findet kein Skill seine Skripte.
    """
    ergebnis = {}
    for quelle_name, ziel_name, endungen in (
        ("scripts", "_system/skripte", (".py",)),
        ("vorlagen", "_system/vorlagen", (".docx", ".pptx", ".md")),
    ):
        quelle = plugin_wurzel / quelle_name
        ziel = wurzel / ziel_name
        ziel.mkdir(parents=True, exist_ok=True)
        kopiert, uebersprungen = [], []
        if not quelle.is_dir():
            ergebnis[quelle_name] = {"fehler": "Quellordner fehlt: %s" % quelle}
            continue
        for datei in sorted(quelle.iterdir()):
            if not datei.is_file() or datei.suffix.lower() not in endungen:
                continue
            # laufzeit_test.py ist Wegwerf-Werkzeug aus Phase 1a und hat im
            # Arbeitsordner der Lehrkraft nichts zu suchen.
            if datei.name == "laufzeit_test.py":
                uebersprungen.append(datei.name)
                continue
            shutil.copy2(str(datei), str(ziel / datei.name))
            kopiert.append(datei.name)
        ergebnis[quelle_name] = {"ziel": str(ziel), "kopiert": kopiert,
                                 "uebersprungen": uebersprungen}
    return ergebnis


# --------------------------------------------------------------------------
# Schritt 3: Index und Prueflisten
# --------------------------------------------------------------------------

def index_anlegen(wurzel):
    ziel = wurzel / "wissensbasis" / "_index.md"
    if ziel.exists():
        return {"pfad": str(ziel), "neu": False}
    ziel.write_text(INDEX_KOPF, encoding="utf-8")
    return {"pfad": str(ziel), "neu": True}


def stand_anlegen(wurzel):
    """_system/stand.json haelt die Pruefwerte fuer die Aenderungserkennung.
    Vorhandene Werte bleiben stehen - sie sind der einzige Beleg dafuer, was
    seit der letzten Ausgabe passiert ist."""
    ziel = wurzel / "_system" / "stand.json"
    if ziel.exists():
        try:
            inhalt = json.loads(ziel.read_text(encoding="utf-8"))
            inhalt["schema"] = SCHEMA
            inhalt["zuletzt_eingerichtet"] = jetzt()
            ziel.write_text(json.dumps(inhalt, indent=2, ensure_ascii=False),
                            encoding="utf-8")
            return {"pfad": str(ziel), "neu": False,
                    "eintraege": len(inhalt.get("dateien", {}))}
        except Exception as fehler:
            return {"pfad": str(ziel), "neu": False,
                    "fehler": "vorhandene Datei unlesbar: %s" % fehler}
    ziel.write_text(json.dumps(
        {"schema": SCHEMA, "zuletzt_eingerichtet": jetzt(), "dateien": {}},
        indent=2, ensure_ascii=False), encoding="utf-8")
    return {"pfad": str(ziel), "neu": True, "eintraege": 0}


# --------------------------------------------------------------------------
# Schritt 4: Migration (Kap. 5.3)
# --------------------------------------------------------------------------

def schema_pruefen(wurzel):
    """Sucht .md-Dateien mit aelterer schema-Version. Version 1 kennt nur
    schema 1 - der Befund ist also bisher immer leer. Die Pruefung steht
    trotzdem hier, damit spaetere Formatwechsel eine Stelle haben."""
    veraltet, ohne_kopf = [], []
    for datei in sorted((wurzel / "wissensbasis").rglob("*.md")):
        if datei.name == "_index.md":
            continue
        try:
            kopf = datei.read_text(encoding="utf-8", errors="replace")[:600]
        except Exception:
            continue
        if not kopf.lstrip().startswith("---"):
            ohne_kopf.append(str(datei.relative_to(wurzel)))
            continue
        for zeile in kopf.splitlines():
            if zeile.startswith("schema:"):
                try:
                    if int(zeile.split(":", 1)[1].strip()) < SCHEMA:
                        veraltet.append(str(datei.relative_to(wurzel)))
                except ValueError:
                    pass
                break
    return {"aktuelles_schema": SCHEMA, "veraltet": veraltet,
            "ohne_metadatenkopf": ohne_kopf}


# --------------------------------------------------------------------------
# Schritt 5: Bereitschaft je Befehl
# --------------------------------------------------------------------------

def vorlage_finden(wurzel):
    """Vorlage der Schule schlaegt Standardvorlage."""
    eigene = wurzel / "_kontext" / "vorlage.docx"
    if eigene.exists():
        return {"art": "schule", "pfad": str(eigene)}
    standard = wurzel / "_system" / "vorlagen" / "standard_vorlage.docx"
    if standard.exists():
        return {"art": "standard", "pfad": str(standard)}
    return {"art": None, "pfad": None}


def bereitschaft(wurzel):
    vorlage = vorlage_finden(wurzel)
    ergebnis = {}
    for befehl, bedingungen in BEDARF.items():
        fehlt = []
        for bedingung in bedingungen:
            if bedingung == "vorlage":
                if not vorlage["art"]:
                    fehlt.append("vorlage")
            elif not (wurzel / bedingung).exists():
                fehlt.append(bedingung)
        ergebnis[befehl] = {
            "bereit": not fehlt,
            "fehlt": fehlt,
            "fehlt_klartext": [KLARTEXT.get(f, f) for f in fehlt],
        }
    ergebnis["_vorlage"] = vorlage
    return ergebnis


def eingang_sichten(wurzel):
    dateien = [d for d in sorted((wurzel / "eingang").iterdir()) if d.is_file()] \
        if (wurzel / "eingang").is_dir() else []
    return {"anzahl": len(dateien), "namen": [d.name for d in dateien[:50]]}


# --------------------------------------------------------------------------
# Bericht
# --------------------------------------------------------------------------

def bericht(daten):
    z = []
    z.append("Einrichtung des Arbeitsordners")
    z.append("=" * 60)
    z.append("Ordner: %s" % daten["arbeitsordner"])
    z.append("")

    o = daten["ordner"]
    if o["angelegt"]:
        z.append("Neu angelegt: %d Ordner (%s)" % (len(o["angelegt"]), ", ".join(o["angelegt"])))
    else:
        z.append("Alle Ordner waren bereits vorhanden.")

    for name, beschriftung in (("scripts", "Skripte"), ("vorlagen", "Vorlagen")):
        teil = daten["bereitgestellt"].get(name, {})
        if teil.get("fehler"):
            z.append("%s: FEHLER – %s" % (beschriftung, teil["fehler"]))
        else:
            z.append("%s bereitgestellt: %d (%s)" % (
                beschriftung, len(teil.get("kopiert", [])),
                ", ".join(teil.get("kopiert", [])) or "keine"))

    z.append("Index: %s" % ("neu angelegt" if daten["index"]["neu"] else "vorhanden"))
    anzahl = daten["stand"].get("eintraege", 0)
    z.append("Prüfwerte: %s, %d %s" % (
        "neu angelegt" if daten["stand"]["neu"] else "fortgeschrieben",
        anzahl, "Eintrag" if anzahl == 1 else "Einträge"))

    migration = daten["migration"]
    if migration["veraltet"]:
        z.append("Migration nötig: %d Dateien mit älterem Format" % len(migration["veraltet"]))
    if migration["ohne_metadatenkopf"]:
        z.append("Ohne Metadaten-Kopf: %d Dateien" % len(migration["ohne_metadatenkopf"]))

    z.append("")
    z.append("Bereitschaft je Befehl")
    z.append("-" * 60)
    bereit = daten["bereitschaft"]
    for befehl in ("aufnehmen", "material", "klassenarbeit"):
        eintrag = bereit[befehl]
        if eintrag["bereit"]:
            z.append("  bereit       /%s" % befehl)
        else:
            z.append("  eingeschränkt /%s – es fehlt: %s"
                     % (befehl, ", ".join(eintrag["fehlt_klartext"])))
    vorlage = bereit["_vorlage"]
    z.append("")
    if vorlage["art"] == "schule":
        z.append("Word-Vorlage: die deiner Schule (_kontext/vorlage.docx)")
    elif vorlage["art"] == "standard":
        z.append("Word-Vorlage: Standardvorlage des Plugins (keine eigene hinterlegt)")
    else:
        z.append("Word-Vorlage: keine gefunden")

    eingang = daten["eingang"]
    if eingang["anzahl"] == 1:
        z.append("")
        z.append("Im Eingang liegt 1 Datei.")
    elif eingang["anzahl"]:
        z.append("")
        z.append("Im Eingang liegen %d Dateien." % eingang["anzahl"])
    return "\n".join(z)


# --------------------------------------------------------------------------
# Hauptprogramm
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Richtet den Arbeitsordner der Lehrkraft ein (Kap. 6.1)")
    parser.add_argument("--ordner", required=True,
                        help="Arbeitsordner der Lehrkraft (z. B. der angehängte Ordner)")
    parser.add_argument("--plugin", default=None,
                        help="Wurzel des Plugins. Standard: zwei Ebenen über diesem Skript")
    parser.add_argument("--json", action="store_true",
                        help="Ergebnis als JSON ausgeben statt als Text")
    argumente = parser.parse_args()

    wurzel = Path(argumente.ordner).expanduser()
    wurzel.mkdir(parents=True, exist_ok=True)
    # Erst anlegen, dann aufloesen: resolve() liefert unter Python < 3.10 auf
    # Windows einen relativen Pfad, wenn der Ordner noch nicht existiert.
    wurzel = wurzel.resolve()

    plugin = Path(argumente.plugin).expanduser().resolve() if argumente.plugin \
        else Path(__file__).resolve().parent.parent

    daten = {
        "schema": SCHEMA,
        "zeitpunkt": jetzt(),
        "arbeitsordner": str(wurzel),
        "pluginwurzel": str(plugin),
    }
    daten["ordner"] = ordner_anlegen(wurzel)
    daten["bereitgestellt"] = bereitstellen(plugin, wurzel)
    daten["index"] = index_anlegen(wurzel)
    daten["stand"] = stand_anlegen(wurzel)
    daten["migration"] = schema_pruefen(wurzel)
    daten["bereitschaft"] = bereitschaft(wurzel)
    daten["eingang"] = eingang_sichten(wurzel)

    if argumente.json:
        print(json.dumps(daten, indent=2, ensure_ascii=False, default=str))
    else:
        print(bericht(daten))
    return 0


if __name__ == "__main__":
    try:
        ENDE = main()
    except BaseException:
        import traceback
        traceback.print_exc()
        ENDE = 1
    sys.exit(ENDE)
