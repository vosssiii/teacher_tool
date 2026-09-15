# Unterrichtsassistent – Phase 1a

Plugin für Lehrkräfte (Englisch, Arbeit & Recht). Konzept: [`../KONZEPT.md`](../KONZEPT.md).

**Stand: Phase 1a (Laufzeit-Test).** Das Plugin enthält bewusst nur einen
Skill und ein Skript. Sie beantworten die Testpunkte T1–T9 aus Kap. 13 und
werden danach durch die echten Skills ersetzt. Es ist Wegwerf-Werkzeug.

## Wozu

Entwickelt wird in Claude Code, ausgeführt in Claude Cowork. Ob dort
Skripte, Bibliotheken und Umwandlungen überhaupt laufen, ist ungeprüft –
das größte Risiko des Vorhabens (Kap. 2). Phase 1a klärt das, **bevor**
weitergebaut wird.

## Was hier liegt

```
unterrichtsassistent/
├── .claude-plugin/plugin.json
├── skills/laufzeit-test/SKILL.md   Ablauf inkl. der fünf Beobachtungen
└── scripts/laufzeit_test.py        misst T1, T2, T3, T6
```

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
- `--install` läuft in zwei Durchgängen: unverzichtbare Bibliotheken und
  getrennt davon die Wege zum PDF. Die scheitern aus eigenen Gründen –
  `weasyprint` etwa installiert sich per pip und startet trotzdem nicht,
  wenn die GTK-Systembibliotheken fehlen. „Installierbar“ und „läuft“ sind
  nicht dasselbe, und genau dieser Unterschied entscheidet T2.

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
