#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Referenztests fuer die Umwandlung (KONZEPT.md Kap. 12, UMWANDLUNG.md).

Jeder Test haelt einen Fall fest, der schon einmal schiefging oder leicht
schiefgehen kann. Nach jeder Aenderung an zu_markdown.py laufen lassen:

    python tests/test_umwandlung.py

Die Dateien unter tests/referenz/ sind selbst gebaut und duerfen ins Repo.
Echtes Material der Lehrkraft liegt - falls vorhanden - in test_datein/ im
Projektstamm und wird nur lokal mitgeprueft; es darf nie ins Repo.

Braucht die Bibliotheken aus der Cowork-Umgebung (python-docx, python-pptx,
pypdf, Pillow). Der Vollstaendigkeitsvergleich braucht zusaetzlich
LibreOffice oder MS Word; ohne beides wird er uebersprungen, nicht verfehlt.
"""

import sys
import tempfile
from pathlib import Path

HIER = Path(__file__).resolve().parent
sys.path.insert(0, str(HIER.parent / "scripts"))
import zu_markdown as zm                 # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

REFERENZ = HIER / "referenz"
ECHT = HIER.parent.parent / "test_datein"

ergebnisse = []


def pruefe(name, bedingung, einzelheit=""):
    ergebnisse.append((name, bool(bedingung), einzelheit))


def umwandeln(pfad):
    ziel = Path(tempfile.mkdtemp(prefix="ua-test-"))
    bericht = zm.umwandeln(pfad, ziel)
    inhalt = (ziel / "inhalt.md").read_text(encoding="utf-8") if (ziel / "inhalt.md").exists() else ""
    return bericht, inhalt


def vollstaendig_ok(bericht):
    """None = Vergleich nicht moeglich (keine Ansicht), sonst True/False."""
    v = bericht.get("vollstaendigkeit")
    if not v or "auffaellig" not in v:
        return None
    return not v["auffaellig"]


# --- Word: Textfelder und automatische Nummerierung ------------------------
bericht, md = umwandeln(REFERENZ / "textfeld_und_nummerierung.docx")
pruefe("Word: ohne Fehler", not bericht.get("fehler"), bericht.get("fehler", ""))
pruefe("Word: automatische Nummer 1.", "1. Nennen Sie" in md)
pruefe("Word: automatische Nummer 3.", "3. Beurteilen Sie" in md)
pruefe("Word: Textfeld als Kasten", "> Merkkasten:" in md)
ok = vollstaendig_ok(bericht)
if ok is not None:
    pruefe("Word: Vollständigkeit ohne Alarm", ok, str(bericht["vollstaendigkeit"].get("beispiele")))

# --- Word: zugeschnittenes Bild ----------------------------------------------
bericht, md = umwandeln(REFERENZ / "zugeschnitten.docx")
pruefe("Word: Zuschnitt erkannt", "In Word zugeschnitten" in md)

# --- PowerPoint: Ebenen und Zuschnitt ----------------------------------------
bericht, md = umwandeln(REFERENZ / "ebenen.pptx")
bildzeilen = [z for z in md.splitlines() if "](bilder/" in z]
pruefe("PowerPoint: drei Bilder", len(bildzeilen) == 3, str(len(bildzeilen)))
pruefe("PowerPoint: Text über Bild A erkannt",
       len(bildzeilen) > 0 and "liegt Text darüber" in bildzeilen[0])
pruefe("PowerPoint: Zuschnitt Bild B erkannt",
       len(bildzeilen) > 1 and "zugeschnitten" in bildzeilen[1])
pruefe("PowerPoint: Bild C ohne Warnung",
       len(bildzeilen) > 2 and "ACHTUNG" not in bildzeilen[2])

# --- Echtes Material, nur lokal ----------------------------------------------
pdf = ECHT / "AB1-Aufsichtspflicht neu gesetzt.pdf"
if pdf.exists():
    bericht, md = umwandeln(pdf)
    ueber = bericht.get("bilder_mit_text_darueber") or {}
    # Fall 1 aus UMWANDLUNG.md: genau die Hausgrafik, keines der Symbole
    pruefe("Echt/PDF: genau ein Bild mit Text darüber", len(ueber) == 1, str(list(ueber)))
    pruefe("Echt/PDF: richtige Buchstaben erkannt",
           any("GEOPSERONENSR" in t for t in ueber.values()), str(list(ueber.values())))
    pruefe("Echt/PDF: Überschrift mit Fragezeichen", "## Was heißt das jetzt für mich" in md)
    ok = vollstaendig_ok(bericht)
    if ok is not None:
        pruefe("Echt/PDF: Vollständigkeit ohne Alarm", ok)

word = ECHT / "UPDATED_Reading Comprehension_VOC WORK.docx"
if word.exists():
    bericht, md = umwandeln(word)
    pruefe("Echt/Word: Unterstreichung erhalten", "<u>eat much less food</u>" in md)
    pruefe("Echt/Word: fetter Aufzählungspunkt", "- restrict – restriction – restrictive" in md)
    pruefe("Echt/Word: Lösungsteil erkannt", bool(bericht.get("loesungsteil_ab")))
    pruefe("Echt/Word: Schreibzeilen zusammengefasst", "[Schreibzeilen: 6]" in md)
    ok = vollstaendig_ok(bericht)
    if ok is not None:
        pruefe("Echt/Word: Vollständigkeit ohne Alarm", ok,
               str(bericht["vollstaendigkeit"].get("beispiele")))

# --- Ergebnis -------------------------------------------------------------
breite = max(len(n) for n, _, _ in ergebnisse)
for name, gut, einzelheit in ergebnisse:
    print("%s  %s%s" % ("ok    " if gut else "FEHLER", name.ljust(breite),
                        ("  – " + einzelheit) if (einzelheit and not gut) else ""))
fehler = sum(1 for _, gut, _ in ergebnisse if not gut)
print("\n%d Prüfungen, %d fehlgeschlagen%s" % (
    len(ergebnisse), fehler, "" if ECHT.exists() else " (echtes Material nicht vorhanden, übersprungen)"))
sys.exit(1 if fehler else 0)
