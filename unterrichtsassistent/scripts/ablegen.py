#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Legt fertig bearbeitete Aufnahmen in die Wissensbasis (KONZEPT.md Kap. 6.2,
Schritte 5 bis 8).

Arbeitet alle Arbeitsbereiche unter _system/aufnahme/ ab, die zu_markdown.py
angelegt und Claude bearbeitet hat. Je Aufnahme:

  1. Pflichtpruefung des Metadatenkopfs - fehlt etwas, wird NICHT abgelegt
     (Designprinzip 3: was immer gelten muss, prueft ein Skript)
  2. Skript-Platzhalter noch da (Bild nicht angesehen, Scan nicht
     abgeschrieben)? -> nicht ablegen
  3. Doppelt? Liegt dasselbe Original schon in originale/ -> nicht ablegen,
     Datei bleibt im Eingang, Claude fragt nach
  4. .md nach wissensbasis/<fach>/<typ>/, einheitlicher Dateiname
  5. verlinkte Bilder nach wissensbasis/<fach>/bilder/, Links anpassen;
     nicht verlinkte Bilder (von Claude als Schmuck gestrichen) verwerfen
  6. Original von eingang/ nach originale/ - verschoben, nie geloescht
  7. Pruefwert in _system/stand.json, Index neu bauen

Nichts wird ueberschrieben: Belegte Dateinamen bekommen -2, -3 angehaengt.

Aufruf:
    python ablegen.py --ordner "<Arbeitsordner>"
    python ablegen.py --ordner "<Arbeitsordner>" --nur <arbeitsbereich>
"""

import argparse
import json
import re
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import gemeinsam as g                      # noqa: E402
import index_aktualisieren                 # noqa: E402

BILDLINK = re.compile(r"(!\[[^\]]*\]\()bilder/([^)\s]+)(\))")


# --------------------------------------------------------------------------
# Pruefung
# --------------------------------------------------------------------------

def pruefen(felder, rumpf):
    """Gibt (fehler, hinweise) zurueck. Fehler blockieren das Ablegen."""
    fehler, hinweise = [], []

    def leer(schluessel):
        return felder.get(schluessel) in (None, "")

    if felder.get("fach") not in g.FAECHER:
        fehler.append("fach fehlt oder ist ungültig (erlaubt: %s)" % ", ".join(g.FAECHER))
    if felder.get("typ") not in g.TYPEN:
        fehler.append("typ fehlt oder ist ungültig (erlaubt: %s)" % ", ".join(g.TYPEN))
    if felder.get("herkunft") not in g.HERKUNFT:
        fehler.append("herkunft fehlt oder ist ungültig (erlaubt: %s)" % ", ".join(g.HERKUNFT))
    if leer("thema"):
        fehler.append("thema fehlt")
    if leer("beschreibung"):
        fehler.append("beschreibung fehlt – ohne sie findet Claude das Material später nicht")
    if felder.get("herkunft") == "fremd" and leer("quelle"):
        fehler.append("quelle fehlt – bei fremdem Material Pflicht (Urheberrecht, Kap. 2.1)")
    if felder.get("typ") == "gesetzestext":
        if leer("url"):
            fehler.append("url fehlt – Gesetzestexte brauchen die offizielle Fundstelle (Kap. 8.2)")
        if leer("rechtsstand"):
            fehler.append("rechtsstand fehlt – bei Gesetzestexten Pflicht (Kap. 8.2)")
    elif felder.get("fach") == "arbeit-recht" and leer("rechtsstand"):
        hinweise.append("rechtsstand nicht angegeben – bei Arbeit & Recht empfohlen")
    if leer("klasse"):
        hinweise.append("klasse nicht angegeben")
    if felder.get("status") not in ("pruefen", "bereit"):
        hinweise.append("status '%s' ist in der Wissensbasis nicht vorgesehen – "
                        "wird 'pruefen'" % felder.get("status"))

    fehler.extend(g.platzhalter_finden(rumpf))
    return fehler, hinweise


# --------------------------------------------------------------------------
# Ablegen
# --------------------------------------------------------------------------

def doppelt(wurzel, pruef, stand):
    for md, eintrag in stand.get("aufgenommen", {}).items():
        if eintrag.get("pruefwert") == pruef:
            return md
    originale = wurzel / "originale"
    if originale.is_dir():
        for datei in originale.rglob("*"):
            if datei.is_file() and g.pruefwert(datei) == pruef:
                return g.relativ(datei, wurzel)
    return None


def ablegen_eine(wurzel, bereich, stand):
    auftrag = json.loads((bereich / "auftrag.json").read_text(encoding="utf-8"))
    ergebnis = {"arbeitsbereich": bereich.name, "quelle": auftrag.get("quelle")}
    felder, rumpf = g.md_lesen(bereich / "inhalt.md")

    fehler, hinweise = pruefen(felder, rumpf)
    ergebnis["hinweise"] = hinweise
    if fehler:
        ergebnis["abgelegt"] = False
        ergebnis["fehler"] = fehler
        return ergebnis

    quelle = wurzel / auftrag["quelle"]
    if not quelle.exists():
        ergebnis["abgelegt"] = False
        ergebnis["fehler"] = ["Original nicht mehr im Eingang: %s" % auftrag["quelle"]]
        return ergebnis

    schon_da = doppelt(wurzel, auftrag["pruefwert"], stand)
    if schon_da:
        ergebnis["abgelegt"] = False
        ergebnis["doppelt"] = schon_da
        ergebnis["fehler"] = ["Dasselbe Original ist schon aufgenommen: %s. "
                              "Die Datei bleibt im Eingang." % schon_da]
        return ergebnis

    # --- Ziel der .md ---------------------------------------------------
    fach, typ = felder["fach"], felder["typ"]
    stamm = g.kurzname(felder["thema"], 50)
    if felder.get("klasse") not in (None, ""):
        stamm += "-kl" + g.kurzname(felder["klasse"], 10)
    ziel_md = g.freier_pfad(wurzel / "wissensbasis" / fach / typ / (stamm + ".md"))
    stamm = ziel_md.stem

    # --- Bilder: nur verlinkte wandern mit ---------------------------------
    bilder_ziel = wurzel / "wissensbasis" / fach / "bilder"
    bilder_ziel.mkdir(parents=True, exist_ok=True)
    zuordnung, verworfen = {}, []
    verlinkt = [m.group(2) for m in BILDLINK.finditer(rumpf)]
    nummer = 0
    for name in verlinkt:
        if name in zuordnung:
            continue
        alt = bereich / "bilder" / name
        if not alt.exists():
            ergebnis.setdefault("hinweise", []).append("verlinktes Bild fehlt: %s" % name)
            continue
        nummer += 1
        neu = g.freier_pfad(bilder_ziel / ("%s-%02d%s" % (stamm, nummer, alt.suffix.lower())))
        shutil.move(str(alt), str(neu))
        zuordnung[name] = neu.name
    if (bereich / "bilder").is_dir():
        verworfen = [d.name for d in (bereich / "bilder").iterdir() if d.is_file()]

    # Links zeigen von wissensbasis/<fach>/<typ>/ nach ../bilder/
    rumpf = BILDLINK.sub(lambda m: "%s../bilder/%s%s" % (
        m.group(1), zuordnung.get(m.group(2), m.group(2)), m.group(3)), rumpf)

    # --- Original verschieben, nie loeschen ---------------------------------
    ziel_original = g.freier_pfad(wurzel / "originale" / quelle.name)
    ziel_original.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(quelle), str(ziel_original))

    # --- Kopf vervollstaendigen und schreiben -------------------------------
    felder["schema"] = g.SCHEMA
    felder["original"] = g.relativ(ziel_original, wurzel)
    felder["grafiken"] = "ja" if zuordnung else "nein"
    felder["aktualisiert"] = g.heute()
    felder.setdefault("erstellt", g.heute())
    if g.PRUEFMARKER.search(rumpf) or felder.get("status") not in ("pruefen", "bereit"):
        # Offene [Prüfen:]- oder [unsicher:]-Stellen: nicht als Quelle nutzen
        felder["status"] = "pruefen"
    g.md_schreiben(ziel_md, felder, rumpf)

    md_rel = g.relativ(ziel_md, wurzel)
    stand.setdefault("aufgenommen", {})[md_rel] = {
        "original": felder["original"],
        "pruefwert": auftrag["pruefwert"],
        "abgelegt": g.jetzt(),
    }
    shutil.rmtree(str(bereich))

    ergebnis.update({
        "abgelegt": True,
        "md": md_rel,
        "original": felder["original"],
        "status": felder["status"],
        "bilder": sorted(zuordnung.values()),
        "bilder_verworfen": verworfen,
    })
    return ergebnis


def alle_ablegen(wurzel, nur=None):
    wurzel = Path(wurzel).resolve()
    aufnahme = wurzel / "_system" / "aufnahme"
    stand = g.stand_lesen(wurzel)
    ergebnisse = []
    bereiche = sorted(p for p in aufnahme.iterdir() if p.is_dir()) if aufnahme.is_dir() else []
    for bereich in bereiche:
        if nur and bereich.name != nur:
            continue
        if not (bereich / "auftrag.json").exists() or not (bereich / "inhalt.md").exists():
            continue
        try:
            ergebnisse.append(ablegen_eine(wurzel, bereich, stand))
        except Exception as fehler:
            ergebnisse.append({"arbeitsbereich": bereich.name, "abgelegt": False,
                               "fehler": ["{}: {}".format(type(fehler).__name__, fehler)]})
        # Nach jeder Datei sichern: bricht der Lauf ab, bleibt der Stand stimmig
        g.stand_schreiben(wurzel, stand)

    index = index_aktualisieren.neu_bauen(wurzel)
    return {
        "abgelegt": [e for e in ergebnisse if e.get("abgelegt")],
        "nicht_abgelegt": [e for e in ergebnisse if not e.get("abgelegt")],
        "index": index,
    }


def text_bericht(daten):
    z = ["Ablage in die Wissensbasis", "=" * 60]
    for e in daten["abgelegt"]:
        z.append("  abgelegt  %s" % e["md"])
        z.append("            Status: %s · Original: %s" % (e["status"], e["original"]))
        if e.get("bilder"):
            z.append("            Bilder: %s" % ", ".join(e["bilder"]))
        if e.get("bilder_verworfen"):
            z.append("            als Schmuck verworfen: %d" % len(e["bilder_verworfen"]))
        for h in e.get("hinweise", []):
            z.append("            Hinweis: %s" % h)
    for e in daten["nicht_abgelegt"]:
        z.append("  NICHT abgelegt  %s" % (e.get("quelle") or e["arbeitsbereich"]))
        for f in e.get("fehler", []):
            z.append("            – %s" % f)
    z.append("")
    z.append("Abgelegt: %d · offen: %d · Index: %d Materialien"
             % (len(daten["abgelegt"]), len(daten["nicht_abgelegt"]),
                daten["index"]["materialien"]))
    return "\n".join(z)


def main():
    parser = argparse.ArgumentParser(description="Bearbeitete Aufnahmen in die Wissensbasis legen")
    parser.add_argument("--ordner", required=True)
    parser.add_argument("--nur", help="Nur diesen Arbeitsbereich ablegen")
    parser.add_argument("--json", action="store_true")
    argumente = parser.parse_args()
    daten = alle_ablegen(argumente.ordner, argumente.nur)
    g.ausgeben(daten, argumente.json, text_bericht)
    return 0


if __name__ == "__main__":
    try:
        ENDE = main()
    except BaseException:
        import traceback
        traceback.print_exc()
        ENDE = 1
    sys.exit(ENDE)
