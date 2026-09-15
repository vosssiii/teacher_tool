# Unterrichtsassistent – Phase 1a

Plugin für Lehrkräfte (Englisch, Arbeit & Recht).

> Das Konzeptpapier (`KONZEPT.md`, Kapitelverweise unten) liegt bewusst
> nicht im Repository. Kapitel- und Testpunktnummern sind hier nur als
> Herkunftsangabe genannt – zum Ausführen des Tests wird es nicht gebraucht.

**Stand: Phase 1b begonnen.** Phase 1a (Laufzeit-Test) ist abgeschlossen –
alle Testpunkte T1–T9 beantwortet, Befunde in `BEFUNDE-PHASE-1A.md`
(liegt nicht im Repo).

| Befehl | Stand |
| --- | --- |
| `/unterrichtsassistent:einrichten` | gebaut |
| `/unterrichtsassistent:aufnehmen` | offen |
| `/unterrichtsassistent:material` | offen |
| `/unterrichtsassistent:klassenarbeit` | offen (braucht Material der Lehrkraft) |
| `/unterrichtsassistent:freigeben` | offen |
| `/unterrichtsassistent:laufzeit-test` | Wegwerf-Werkzeug aus Phase 1a |

## Was hier liegt

```
unterrichtsassistent/
├── .claude-plugin/plugin.json
├── commands/                       ohne diese Dateien gibt es keine Befehle
│   ├── einrichten.md
│   └── laufzeit-test.md
├── skills/
│   ├── einrichten/SKILL.md         Ordner vorbereiten, Kontext erfassen
│   └── laufzeit-test/SKILL.md      Ablauf des Laufzeit-Tests
├── scripts/
│   ├── einrichten.py               Ordner, Skript-Bereitstellung, Bereitschaft
│   ├── vorlage_pruefen.py          Formatvorlagen einer Vorlage prüfen
│   └── laufzeit_test.py            misst T1, T2, T3, T6
└── vorlagen/
    ├── standard_vorlage.docx       neutrale Word-Vorlage als Ersatz
    ├── standard_vorlage.pptx       neutrale PowerPoint-Vorlage
    ├── schule_vorlage.md           Aufbau von _kontext/schule.md
    └── ordner_anweisung.md         Text für die Anweisung im Projekt
```

## Zwei Dinge, die aus Phase 1a folgen

**Befehle brauchen `commands/`.** Ein Skill allein ist kein
Schrägstrich-Befehl. Cowork listet ihn, meldet beim Aufruf aber „Unknown
command". Jeder Befehl braucht eine Datei in `commands/`, die den Skill
aufruft. Das Präfix gehört dazu: `/unterrichtsassistent:einrichten`.

**Skripte müssen an ihren Einsatzort gebracht werden.**
`${CLAUDE_PLUGIN_ROOT}` ist leer, und in der lokalen Geräte-Sandbox liegt
das Plugin gar nicht. `einrichten.py` kopiert deshalb alle Skripte und
Vorlagen nach `_system/skripte/` bzw. `_system/vorlagen/` im Arbeitsordner
der Lehrkraft und frischt sie bei jedem Aufruf auf. Skripte nehmen keine
festen Pfade an – der Arbeitsordner wird übergeben.

### Warum es beides gibt, `commands/` und `skills/`

Befund aus dem ersten Cowork-Lauf: Ein Skill unter `skills/*/SKILL.md` wird
zwar geladen und kann vom Modell aufgerufen werden, ist aber **kein**
Schrägstrich-Befehl. Cowork zeigt ihn in der Liste an, fügt beim Auswählen
`/unterrichtsassistent:laufzeit-test` ein – und antwortet darauf
„Unknown command“. Nur Dateien in `commands/` werden zu Befehlen.

Das gilt für jeden Befehl, den die Lehrkraft tippen soll. `/einrichten`,
`/material`, `/aufnehmen`, `/klassenarbeit` und `/freigeben` brauchen je
eine Datei in `commands/`, die den zugehörigen Skill aufruft. Die
Plugin-Struktur in Kap. 4.2 des Konzepts sieht das noch nicht vor.

## Die Arbeitsteilung

Vier Testpunkte lassen sich messen, fünf nur beobachten:

| Skript | Beobachtung durch Claude in der Sitzung |
| --- | --- |
| T1 Bibliotheken | T4 Befehlsnamen und Skriptaufruf |
| T2 Word → PDF, Ersatzweg über HTML | T5 Modellwahl |
| T3 Word/PowerPoint aus Vorlage | T7 Ordner-Anweisung |
| T6 Webzugriff im Skript | T8 freie Formulierungen |
| | T9 Dateien pro Durchlauf |

Beides landet in einem Bericht. Solange dort „offen“ steht, ist Phase 1a
nicht abgeschlossen.

## Skript allein aufrufen

```bash
python scripts/laufzeit_test.py --umgebung "Claude Cowork"
python scripts/laufzeit_test.py --umgebung "Claude Cowork" --install
```

| Schalter | Wirkung |
| --- | --- |
| `--umgebung TEXT` | Wo der Lauf stattfand. Steht im Bericht – ein Lauf in Claude Code sagt nichts über Cowork aus. |
| `--ausgabe PFAD` | Zielordner, Standard `./laufzeit-test` |
| `--install` | Rüstet fehlende Bibliotheken nach. Ohne den Schalter wird nichts verändert. |

Ergebnis in `laufzeit-test/`: `bericht.md` (lesbar) und `roh.json`
(Rohdaten, um Läufe zu vergleichen).

**Grundsätze des Skripts**

- Reine Standardbibliothek. Es muss starten, bevor klar ist, ob überhaupt
  etwas installiert ist.
- Kein Probenfehler bricht den Lauf ab, Exit-Code immer 0. Ein Fehlschlag
  ist ein Ergebnis, kein Absturz.
- Eine erzeugte Datei beweist nichts. Word und PowerPoint werden
  **zurückgelesen** und geprüft: Wurden die Formatvorlagen angewendet, ist
  die Tabelle erhalten, sind Umlaute und `§` intakt?
- `--install` läuft in zwei Durchgängen: unverzichtbare Bibliotheken als
  Gruppe, die Wege zum PDF einzeln – pip installiert sonst alles oder
  nichts, und ein Paket, das sich nicht auflösen lässt, würde die anderen
  mitreißen.
- „Installierbar“ und „läuft“ sind zweierlei, und der Unterschied
  entscheidet T2: `weasyprint` installiert sich per pip anstandslos und
  startet trotzdem nicht, wenn die GTK-Systembibliotheken fehlen.
- **Nach `--install` einmal erneut ohne den Schalter laufen lassen.**
  Frisch installierte Pakete sind teils erst in einem neuen Prozess voll
  nutzbar – `docx2pdf` meldet im Installationslauf einen Fehler und
  erzeugt im nächsten ein einwandfreies PDF. Es zählt der letzte Lauf.

## Der Weg zum PDF

T2 prüft drei Kategorien, absteigend nach dem, was die Lehrkraft davon hat:

| Kategorie | Wege | Braucht |
| --- | --- | --- |
| Word → PDF | docx2pdf, LibreOffice, pandoc | MS Word, LibreOffice oder LaTeX |
| .md → HTML → PDF | **xhtml2pdf**, weasyprint, wkhtmltopdf | xhtml2pdf: nichts. Die anderen: GTK bzw. externes Programm |
| PDF direkt bauen | **fpdf2**, **reportlab** | nichts |

Die fett gesetzten sind reine pip-Pakete. Damit ist die PDF-Ausgabe **nicht**
davon abhängig, dass auf dem Rechner der Lehrkraft LibreOffice installiert
ist.

Was davon abhängig bleibt: **nur der Hauptweg liefert ein PDF, das der
Word-Vorlage der Schule folgt.** Kein reines Python-Paket rendert ein
`.docx`-Layout. Fällt der Hauptweg aus, bekommt die Lehrkraft weiterhin
beides – Word-Datei *und* PDF –, aber das PDF sieht anders aus als die
Word-Datei. Kap. 5.2 sieht diesen Ersatzweg bereits vor; der Test sagt, ob
er gebraucht wird.

Zusätzlich wird geprüft, ob der Inhalt im PDF als **Text** ankommt und
nicht als Pixel – ein Arbeitsblatt, das sich nicht durchsuchen, kopieren
oder vorlesen lässt, wäre kein brauchbares Ergebnis.

## In Cowork testen

Bis zum ersten Push lokal einbinden. Über den Marketplace, sobald das Repo
auf GitHub liegt:

```
/plugin marketplace add <benutzer>/teacher_tool
/plugin install unterrichtsassistent@teacher_tool
```

Das Repo muss dafür öffentlich sein, sonst braucht die Lehrkraft einen
GitHub-Zugang – das widerspricht dem Ziel „installiert allein, ohne Hilfe“
(Kap. 6.1). Im Repo liegt nur Plugin-Code. Die Materialien der Lehrkraft
liegen in ihrem Ordner `Unterricht/` und kommen nie hier hinein.

Dann in Cowork den Skill auslösen und den Anweisungen in `SKILL.md` folgen.

## Wann Phase 1a abgeschlossen ist

Wenn alle neun Punkte im Bericht beantwortet sind und das Fazit steht.
Laut Kap. 11 müssen Architektur und Designprinzip 3 („Skript vor
Sprachmodell“) **neu bewertet werden, bevor weitergebaut wird**, falls:

- sich Skripte gar nicht ausführen lassen,
- weder Word noch PowerPoint erzeugt werden können,
- kein Weg zu einem PDF führt,
- der Skill bei keiner freien Formulierung greift.
