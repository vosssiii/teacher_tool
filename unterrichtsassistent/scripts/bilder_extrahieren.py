#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Holt Bilder aus Word-, PDF- und PowerPoint-Dateien (KONZEPT.md Kap. 5.1).

Das Skript holt die Bilder nur heraus. Was darauf zu sehen ist und ob ein
Bild Inhalt traegt oder nur Schmuck ist, entscheidet Claude beim Ansehen -
das kann kein Skript. Deshalb wird hier nichts nach Bedeutung aussortiert,
nur nach Technik: winzige Bilder (Aufzaehlungspunkte, Linien) und doppelte
(dasselbe Logo auf jeder Seite) fallen weg.

Wird von zu_markdown.py als Modul benutzt, laeuft aber auch allein:
    python bilder_extrahieren.py --datei blatt.pdf --ziel bilder/
"""

import argparse
import hashlib
import json
import logging
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# Kleiner als das in beide Richtungen: Aufzaehlungszeichen, Trennlinien.
MINDESTKANTE = 32

ENDUNG_NACH_TYP = {
    "image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif",
    "image/bmp": ".bmp", "image/tiff": ".tif", "image/x-emf": ".emf",
    "image/x-wmf": ".wmf", "image/svg+xml": ".svg", "image/webp": ".webp",
}


class Sammler:
    """Speichert Bilder, verhindert Doppelte, merkt sich, was wegfiel."""

    def __init__(self, ziel, praefix="bild"):
        self.ziel = Path(ziel)
        self.ziel.mkdir(parents=True, exist_ok=True)
        self.praefix = praefix
        self.gesehen = {}          # hash -> dateiname
        self.gespeichert = []
        self.weggelassen = []
        self.zaehler = 0

    def ablegen(self, daten, endung, herkunft):
        """Gibt den Dateinamen zurueck oder None, wenn das Bild wegfaellt."""
        pruef = hashlib.sha256(daten).hexdigest()
        if pruef in self.gesehen:
            return self.gesehen[pruef]
        groesse = _masse(daten)
        if groesse and min(groesse) < MINDESTKANTE:
            self.weggelassen.append({"herkunft": herkunft, "grund": "zu klein",
                                     "masse": "%sx%s" % groesse})
            return None
        self.zaehler += 1
        name = "%s-%02d%s" % (self.praefix, self.zaehler, endung.lower())
        (self.ziel / name).write_bytes(daten)
        self.gesehen[pruef] = name
        self.gespeichert.append({
            "datei": name, "herkunft": herkunft, "bytes": len(daten),
            "masse": ("%sx%s" % groesse) if groesse else None,
            # Vektorformate aus Office kann Claude nicht ansehen.
            "ansehbar": endung.lower() in (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"),
        })
        return name

    def bericht(self):
        return {"gespeichert": self.gespeichert, "weggelassen": self.weggelassen}


def _masse(daten):
    try:
        from io import BytesIO
        from PIL import Image
        with Image.open(BytesIO(daten)) as bild:
            return bild.size
    except Exception:
        return None


# --------------------------------------------------------------------------
# Word
# --------------------------------------------------------------------------

def aus_word_bild(dokument, rid, sammler):
    """Ein einzelnes Bild aus einem Word-Dokument, ueber seine Beziehungs-ID.
    zu_markdown.py ruft das an der Stelle auf, an der das Bild im Text steht -
    so bleibt die Position erhalten."""
    try:
        teil = dokument.part.related_parts[rid]
    except KeyError:
        return None
    endung = ENDUNG_NACH_TYP.get(teil.content_type) or Path(str(teil.partname)).suffix or ".bin"
    return sammler.ablegen(teil.blob, endung, "Word")


def aus_word(pfad, sammler):
    import docx
    dokument = docx.Document(str(pfad))
    for rid, teil in dokument.part.related_parts.items():
        if "image" in teil.content_type:
            aus_word_bild(dokument, rid, sammler)
    return sammler.bericht()


# --------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------

def aus_pdf_seite(seite, seitennummer, sammler):
    namen = []
    try:
        bilder = list(seite.images)
    except Exception as fehler:
        sammler.weggelassen.append({"herkunft": "Seite %d" % seitennummer,
                                    "grund": "nicht lesbar: %s" % fehler})
        return namen
    for bild in bilder:
        try:
            daten = bild.data
            endung = Path(bild.name).suffix or ".png"
        except Exception as fehler:
            sammler.weggelassen.append({"herkunft": "Seite %d" % seitennummer,
                                        "grund": "nicht lesbar: %s" % fehler})
            continue
        name = sammler.ablegen(daten, endung, "Seite %d" % seitennummer)
        if name and name not in namen:
            namen.append(name)
    return namen


def aus_pdf(pfad, sammler):
    import pypdf
    logging.getLogger("pypdf").setLevel(logging.ERROR)
    leser = pypdf.PdfReader(str(pfad))
    for nummer, seite in enumerate(leser.pages, start=1):
        aus_pdf_seite(seite, nummer, sammler)
    return sammler.bericht()


# --------------------------------------------------------------------------
# PowerPoint
# --------------------------------------------------------------------------

def aus_pptx_form(form, foliennummer, sammler):
    try:
        bild = form.image
    except Exception:
        return None
    endung = "." + (bild.ext or "png")
    return sammler.ablegen(bild.blob, endung, "Folie %d" % foliennummer)


def aus_pptx(pfad, sammler):
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    praesentation = Presentation(str(pfad))

    def durchlaufen(formen, nummer):
        for form in formen:
            if form.shape_type == MSO_SHAPE_TYPE.GROUP:
                durchlaufen(form.shapes, nummer)
            elif form.shape_type == MSO_SHAPE_TYPE.PICTURE:
                aus_pptx_form(form, nummer, sammler)

    for nummer, folie in enumerate(praesentation.slides, start=1):
        durchlaufen(folie.shapes, nummer)
    return sammler.bericht()


def main():
    parser = argparse.ArgumentParser(description="Bilder aus Word, PDF oder PowerPoint holen")
    parser.add_argument("--datei", required=True)
    parser.add_argument("--ziel", required=True)
    parser.add_argument("--praefix", default="bild")
    argumente = parser.parse_args()

    pfad = Path(argumente.datei)
    sammler = Sammler(argumente.ziel, argumente.praefix)
    endung = pfad.suffix.lower()
    if endung == ".docx":
        bericht = aus_word(pfad, sammler)
    elif endung == ".pdf":
        bericht = aus_pdf(pfad, sammler)
    elif endung == ".pptx":
        bericht = aus_pptx(pfad, sammler)
    else:
        bericht = {"fehler": "Keine Bilder aus %s-Dateien" % endung}
    print(json.dumps(bericht, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
