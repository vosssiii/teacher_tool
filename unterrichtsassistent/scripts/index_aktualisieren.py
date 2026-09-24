#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Baut wissensbasis/_index.md aus den Metadatenkoepfen neu (KONZEPT.md Kap. 5.5).

Der Index wird jedes Mal vollstaendig neu geschrieben, nie fortgeschrieben.
So kann er nicht auseinanderlaufen: Was hier steht, steht auch im Kopf der
Datei. Wer etwas aendern will, aendert den Kopf.

Claude liest zuerst diesen Index und oeffnet dann gezielt einzelne Dateien -
deshalb muss die Beschreibung aussagekraeftig sein.

Aufruf:
    python index_aktualisieren.py --ordner "<Arbeitsordner>"
"""

import argparse
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import gemeinsam as g                      # noqa: E402

INDEX_KOPF = """# Inhaltsverzeichnis der Wissensbasis

Diese Datei pflegt das Plugin. Bitte nicht von Hand bearbeiten – Änderungen
gehen beim nächsten Durchlauf verloren. Was hier steht, stammt aus dem
Metadaten-Kopf der einzelnen Materialien. Wer etwas ändern will, ändert den
Kopf der jeweiligen Datei.
"""

SPALTEN = ["Datei", "Fach", "Klasse", "Typ", "Thema", "Status", "Herkunft", "Beschreibung"]


def zelle(wert):
    text = "" if wert is None else str(wert)
    return text.replace("|", "\\|").replace("\n", " ").strip()


def eintraege(wurzel):
    basis = wurzel / "wissensbasis"
    ergebnis = []
    for pfad in sorted(basis.rglob("*.md")):
        if pfad.name == "_index.md" or "bilder" in pfad.relative_to(basis).parts:
            continue
        try:
            felder, _ = g.md_lesen(pfad)
        except Exception as fehler:
            felder = {"beschreibung": "(nicht lesbar: %s)" % fehler}
        ergebnis.append((pfad.relative_to(basis).as_posix(), felder))
    return ergebnis


def neu_bauen(wurzel):
    wurzel = Path(wurzel).resolve()
    liste = eintraege(wurzel)
    liste.sort(key=lambda e: (str(e[1].get("fach") or "~"), str(e[1].get("typ") or "~"),
                              str(e[1].get("thema") or "~").lower(), e[0]))

    stati = Counter(str(f.get("status") or "ohne") for _, f in liste)
    ohne_kopf = [datei for datei, f in liste if not f]

    zeilen = [INDEX_KOPF]
    zeilen.append("Stand: %s · %d Materialien (%s)\n" % (
        g.heute(), len(liste),
        ", ".join("%d %s" % (n, s) for s, n in sorted(stati.items())) or "keine"))
    zeilen.append("| " + " | ".join(SPALTEN) + " |")
    zeilen.append("| " + " | ".join("---" for _ in SPALTEN) + " |")
    for datei, f in liste:
        zeilen.append("| " + " | ".join([
            "[%s](%s)" % (zelle(Path(datei).name), datei.replace(" ", "%20")),
            zelle(f.get("fach")), zelle(f.get("klasse")), zelle(f.get("typ")),
            zelle(f.get("thema")), zelle(f.get("status")), zelle(f.get("herkunft")),
            zelle(f.get("beschreibung") if f else "(ohne Metadaten-Kopf)"),
        ]) + " |")

    ziel = wurzel / "wissensbasis" / "_index.md"
    ziel.parent.mkdir(parents=True, exist_ok=True)
    ziel.write_text("\n".join(zeilen) + "\n", encoding="utf-8")
    return {"pfad": str(ziel), "materialien": len(liste), "status": dict(stati),
            "ohne_metadatenkopf": ohne_kopf}


def main():
    parser = argparse.ArgumentParser(description="Inhaltsverzeichnis der Wissensbasis neu bauen")
    parser.add_argument("--ordner", required=True)
    parser.add_argument("--json", action="store_true")
    argumente = parser.parse_args()
    daten = neu_bauen(argumente.ordner)
    g.ausgeben(daten, argumente.json, lambda d: "Index neu gebaut: %d Materialien %s" % (
        d["materialien"], d["status"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
