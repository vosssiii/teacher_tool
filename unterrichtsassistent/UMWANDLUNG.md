# Umwandlung von Unterrichtsmaterial

Wie `/aufnehmen` aus einer Datei der Lehrkraft eine Markdown-Abschrift für
die Wissensbasis macht – je Dateiformat, und wer welchen Teil übernimmt:
ein Skript oder Claude.

Stand: 24.09.2026 · gehört zu `KONZEPT.md` Kap. 5.1 und 6.2

**Zeichen in diesem Dokument:** ✅ umgesetzt und getestet · 🔧 geplant ·
⚠️ bekannte Grenze · ❓ noch nicht mit echtem Material getestet

---

## 1. Worum es geht

Die Umwandlung ist der **fehleranfälligste Schritt des ganzen Plugins**.
Alles, was später entsteht – Arbeitsblätter, Präsentationen,
Klassenarbeiten –, baut auf dem auf, was hier in der Wissensbasis landet.
Ein Fehler in der Abschrift wird zum Fehler im neuen Material.

### Was die Abschrift ist – und was nicht

- **Sie ist eine inhaltliche Abschrift.** Text, Aufgaben, Tabellen, Lösungen
  und die Bilder, die Inhalt tragen.
- **Sie ist kein Layout-Nachbau.** Spalten, Schriftarten, Abstände gehen
  verloren. Das ist Absicht: Das Original bleibt unverändert in
  `originale/` erhalten.

### Vier Grundsätze

1. **Skript für das, was maschinenlesbar in der Datei steckt.** Text aus
   Word, Überschriften aus Formatvorlagen, Tabellen – das liest ein Skript
   zuverlässiger und immer gleich (Designprinzip 3).
2. **Claude für alles, wo man hinsehen muss.** Scans, Fotos, Bilder,
   Einordnung, Urteil.
3. **Die Wahrheit ist, was die Schüler sehen.** Nicht das, was technisch in
   der Datei steckt. Beides kann sich unterscheiden (siehe Kap. 6, Fall 1).
4. **Abschreiben, nicht verbessern.** Fehler im Original bleiben stehen und
   bekommen `[unsicher: …]`. Die Lehrkraft entscheidet.

---

## 2. Der Ablauf

```
eingang/                 _system/aufnahme/<name>/               wissensbasis/<fach>/<typ>/
  Datei  ──zu_markdown.py──▶  inhalt.md     ──Claude──▶  ──ablegen.py──▶  <thema>.md
                              ansicht.pdf    prüft,                        ../bilder/
                              bilder/        ergänzt                     originale/
                              auftrag.json                               _index.md
                                                                         _index.md
```

| Station | Wer | Was |
|---|---|---|
| 1. Umwandeln | Skript `zu_markdown.py` | Datei → Rohabschrift mit leerem Metadatenkopf, Bilder herauslösen, Seitenansicht `ansicht.pdf`, Vollständigkeitsvergleich, Warnvermerke |
| 2. Durchsehen | **Claude** | **Zuerst die Seitenansicht**, dann einordnen, Bilder ansehen, Scans lesen, Unsicheres markieren |
| 3. Ablegen | Skript `ablegen.py` | Pflichtprüfung, Ablage, Original archivieren – **lehnt Unvollständiges ab** |
| 4. Index | Skript `index_aktualisieren.py` | `_index.md` neu aus allen Metadatenköpfen |

Pro Durchlauf höchstens 10 Dateien, damit das Kontingent der Lehrkraft
nicht auf einmal verbraucht wird.

---

## 3. Je Dateiformat

### Word (`.docx`)

| Skript | Claude |
|---|---|
| ✅ Absätze in der Reihenfolge des Dokuments, samt Tabellen | Fach, Thema, Typ, Herkunft, Beschreibung |
| ✅ Überschriften aus Formatvorlagen – auch aus schuleigenen wie `VTitle`, `VHead` | Herausgelöste Bilder ansehen: beschreiben oder als Schmuck streichen |
| ✅ **fett**, *kursiv*, <u>unterstrichen</u>, Links | Lösungsteil mit `[Lösung]` … `[Ende Lösung]` einschließen |
| ✅ Einzelzellen-Tabellen als Kasten (`>`) | Fehler im Original markieren |
| ✅ Schreiblinien → `[Schreibzeilen: n]` | |
| ✅ Lösungsteil erkennen („Teacher's Key", „Lösung" …) | |
| ✅ Bilder an der Stelle, wo sie im Text stehen | |
| ✅ Kopf- und Fußzeilen im Bericht (Hinweis auf Schule, Klasse) | |

**Warum nicht erst in PDF umwandeln?** Naheliegend, weil das PDF zeigt, was
die Schüler sehen. Aber die Word-Datei sagt ausdrücklich, was eine
Überschrift, eine Tabelle, ein unterstrichenes Wort ist. Im PDF stehen nur
Buchstaben an Positionen: Tabellen zerfallen zu losen Zeilen,
Unterstreichungen verschwinden (sie sind im PDF gezeichnete Linien),
Überschriften müssen geraten werden. **Entscheidung:** Umwandlung aus Word,
das PDF zusätzlich als Ansicht und zur Vollständigkeitsprüfung (Kap. 6).

**Grenzen:**

- ✅ **Textfelder** werden gelesen und als Kasten (`>`) übernommen. Word
  speichert sie oft doppelt (moderne Form und Rückfall für alte Programme) –
  genommen wird nur eine.
- ✅ **Automatische Nummerierung** wird nachgerechnet: „1.", „a)", „ii."
  wie in Word angezeigt, auch wenn die Nummerierung aus der Formatvorlage
  kommt.
- ✅ **Zugeschnittene Bilder** bekommen einen `ACHTUNG`-Vermerk: Word
  speichert das ganze Bild und zeigt nur einen Ausschnitt; herausgelöst wird
  das ganze.
- ⚠️ **Formen, die frei über einem Bild liegen**, erkennt das Skript nicht
  als Überdeckung. Ihr Text fällt aber im Vollständigkeitsvergleich auf, und
  die Seitenansicht zeigt die Lage.
- ⚠️ Fußnoten, Kommentare, Formeln, SmartArt und Diagramme werden nicht
  gelesen – der Vollständigkeitsvergleich meldet ihren Text als fehlend.
- ⚠️ Vektorgrafiken (`.emf`, `.wmf`) kann Claude nicht ansehen.

### PDF mit Text

| Skript | Claude |
|---|---|
| ✅ Text aus der Textebene, Seite für Seite, mit `<!-- Seite n -->` | wie bei Word |
| ✅ Zeilen zu Absätzen zusammensetzen, Silbentrennung aufheben | **Überschriften und Absätze prüfen** – hier rät das Skript |
| ✅ Überschriften aus Schriftgröße und Fettdruck | Bilder an die richtige Stelle setzen |
| ✅ Kursiver oder fetter Vorspann als Ganzes | Einzelteile gegen die Seitenansicht prüfen |
| ✅ Bilder herauslösen – winzige und doppelte fallen weg | |

**Grenzen:**

- ⚠️ **PDF kennt keine Überschriften**, nur Schriften. Ist eine
  Zwischenüberschrift weder größer noch fett, erkennt das Skript sie über
  eine Regel (kurze Zeile ohne Satzzeichen, danach langer Absatz). Die kann
  danebenliegen.
- ⚠️ **Die Position von Bildern ist unbekannt.** Sie landen am Seitenende
  mit Vermerk; Claude setzt sie an die richtige Stelle.
- ✅ **Herausgelöste Bilder können vom Sichtbaren abweichen.** Liegt Text
  über einem Bild, zeigt das herausgelöste Bild womöglich einen alten,
  verdeckten Zustand (siehe Kap. 6, Fall 1). Das Skript erkennt das und
  setzt einen `ACHTUNG`-Vermerk samt dem darüberliegenden Text.
- ❓ **Mehrspaltiger Satz:** Die Textreihenfolge kann durcheinandergeraten.
- ⚠️ **Scans mit schlechter Texterkennung:** Hat ein gescanntes PDF bereits
  eine (fehlerhafte) Textebene, hält das Skript es für ein Text-PDF und
  übernimmt den Unsinn. Automatisch erkennt das nichts – nur Claudes Blick
  auf die Seitenansicht.

### PDF-Scan (keine Textebene)

| Skript | Claude |
|---|---|
| ✅ Erkennt Seiten ohne Text (unter 40 Zeichen) | **Liest die Seite als Bild und schreibt sie ab** |
| ✅ Setzt dort einen Platzhalter, der das Ablegen sperrt | Überschriften, Aufzählungen, Tabellen übernehmen |
| ✅ Löst das Seitenbild heraus | Unsicher Lesbares: `[unsicher: vermutlich „…"]` |
| | Handschrift abschreiben, wenn lesbar |

❓ Noch nicht mit einem echten Scan getestet.

### PowerPoint (`.pptx`)

| Skript | Claude |
|---|---|
| ✅ Je Folie eine Überschrift `## Folie n: Titel` | wie bei Word |
| ✅ Textfelder als Stichpunkte, mit Einrückung | Reihenfolge der Inhalte prüfen |
| ✅ Tabellen, Bilder, Sprechernotizen | |
| ✅ Reihenfolge nach Position auf der Folie (oben vor unten, links vor rechts) | |

**Grenzen:** ✅ Folien bestehen aus übereinanderliegenden Ebenen – liegt
Text über einem Bild, bekommt es einen `ACHTUNG`-Vermerk. ✅ Zugeschnittene
Bilder ebenso. ⚠️ SmartArt und Diagramme werden nicht gelesen; der
Vollständigkeitsvergleich meldet ihren Text. ❓ Noch nicht mit echtem
Material getestet.

### Excel (`.xlsx`)

| Skript | Claude |
|---|---|
| ✅ Je Tabellenblatt eine Markdown-Tabelle | Einordnen |
| ✅ Berechnete Werte statt Formeln | |

**Grenzen:** ⚠️ Höchstens 200 Zeilen und 30 Spalten je Blatt, darüber mit
Vermerk gekürzt. ⚠️ Verbundene Zellen werden nicht zusammengefasst.
❓ Noch nicht mit echtem Material getestet.

### Bilddateien (Foto, Infografik, Tafelbild, Screenshot)

| Skript | Claude |
|---|---|
| ✅ Bild unverändert nach `bilder/` | **Beschreibung:** was das Bild zeigt, 2–4 Sätze |
| ✅ Vorlage mit zwei Abschnitten, die das Ablegen sperren, bis sie ausgefüllt sind | **Text im Bild: vollständig abschreiben**, in der Gliederung des Bildes |
| | Einordnen, Herkunft klären |

Die vollständige Abschrift ist Pflicht, eine Zusammenfassung reicht nicht:
Was später in neues Material einfließt, sind die Zahlen und Sätze selbst.

**Grenzen:** ✅ Handyfotos, die quer gespeichert und nur per Vermerk
aufrecht angezeigt werden, richtet das Skript beim Kopieren gerade – sonst
sähe Claude sie womöglich seitlich. ⚠️ `.heic` (iPhone) kann Claude nicht
ansehen; in den Einstellungen des iPhones „Maximale Kompatibilität" wählen
oder als JPG exportieren.

Hier gibt es kein Ebenen-Problem: Was in der Datei steckt, ist, was man
sieht.

### Text und Markdown (`.txt`, `.md`)

✅ Wird übernommen, Zeichenkodierung automatisch erkannt. Hat eine `.md`
schon einen Metadatenkopf, werden dessen Werte übernommen.

### Ältere und fremde Formate

`.doc`, `.odt`, `.rtf`, `.ppt`, `.odp`, `.xls`, `.ods` – ✅ werden zuerst
mit LibreOffice ins moderne Office-Format gebracht, dann wie oben.

❓ **Apple-Formate** (`.pages`, `.key`, `.numbers`) versucht das Skript
ebenfalls über LibreOffice – das ist ungetestet und kann scheitern. Da die
Lehrkraft einen Mac nutzt, ist das wahrscheinlich. Im Zweifel: in Pages als
Word exportieren.

---

## 4. Bilder

| | Eigenständige Bilddatei | Bild in Word, PDF, PowerPoint |
|---|---|---|
| Eigene `.md` | ja | nein – Beschreibung in der `.md` des Dokuments |
| Claude sieht es an | ja | ja |
| Text im Bild | vollständig abschreiben | in die Beschreibung aufnehmen |
| Schmuck | – | Zeile streichen, Bild wird verworfen |

**Inhalt oder Schmuck?** Das entscheidet nur Claude. Das Skript sortiert nur
technisch aus: winzige Bilder (Aufzählungspunkte, Linien) und doppelte
(dasselbe Logo auf jeder Seite). Ergibt sich eine Aufgabe erst aus mehreren
Schmuckelementen zusammen – Symbole mit Pfeilen als Schema –, beschreibt
Claude das Schema in einem Satz und streicht die Einzelbilder.

---

## 5. Was die Abschrift enthält

| Element | Schreibweise |
|---|---|
| Metadatenkopf | YAML zwischen `---` (Kap. 5.3 des Konzepts) |
| Überschriften | `#`, `##`, `###` |
| Schreiblinien | `[Schreibzeilen: n]` |
| Unterstreichung | `<u>…</u>` – Aufgaben verweisen oft darauf |
| Zeilenumbruch im Absatz | bleibt ein Zeilenumbruch (weicht von Standard-Markdown ab) |
| Kasten | `> …` |
| Lösungsteil | `[Lösung]` … `[Ende Lösung]` |
| Seitenwechsel im Original | `<!-- Seite n -->` |
| Unsicheres | `[unsicher: …]` – erzwingt Status `pruefen` |
| Offene Stelle für die Lehrkraft | `[Prüfen: …]` – erzwingt Status `pruefen` |

---

## 6. Bekannte Fehlerquellen

### Fall 1: Herausgelöstes Bild weicht vom Sichtbaren ab

**Was passiert ist** (Testlauf 24.09.2026): Ein PDF-Arbeitsblatt enthält eine
Hausgrafik mit einem Buchstabenrätsel. Die Textebene lieferte richtig
„GEOPSERONENSR" (= Personensorge). Das herausgelöste Bild zeigte dagegen
„GEROPSNENSO" – eine ältere, fehlerhafte Fassung. Auf der Seite ist diese
vermutlich überdeckt; die Schüler sehen die richtige.

Claude hielt das Bild für die Wahrheit und meldete einen Fehler im
Arbeitsblatt, den es nicht gibt.

**Ursache:** Das Skript liefert Text und Bilder getrennt. Was im PDF
*übereinander* liegt, sieht niemand im Zusammenhang.

**Gegenmaßnahmen:**

| | Maßnahme |
|---|---|
| ✅ | **Ansichts-PDF zu jeder Aufnahme.** Bei PDFs das Original, bei Word und PowerPoint per LibreOffice erzeugt. Claude sieht sich die Seiten an. |
| ✅ | **Regel im Skill:** Widersprechen sich Text und Bild, entscheidet die Seitenansicht. |
| ✅ | **Fehler im Original nur melden, wenn sie in der Seitenansicht zu sehen sind.** Ohne diese Prüfung eine Frage stellen, keinen Befund. |
| ✅ | **Warnvermerk des Skripts** an Bildern, über denen etwas liegt oder die zugeschnitten sind: „Bild allein nicht verlässlich". |
| ✅ | **Automatischer Vollständigkeitsvergleich:** Das Skript vergleicht die Wörter der Ansichts-PDF mit denen der Abschrift. Fehlt sichtbarer Text in der Abschrift, meldet es die Wörter. Fängt Textfelder, automatische Nummern, Fußnoten und SmartArt ab, ohne dass jede Lücke einzeln bekannt sein muss. Gilt für Word und PowerPoint. |

❓ Offen: Kann Claude in Cowork PDF-Seiten als Bild ansehen? Muss getestet
werden, bevor die Maßnahmen darauf bauen.

### Weitere Fehlerquellen im Überblick

| Fehlerquelle | Formate | Gegenmaßnahme | Stand |
|---|---|---|---|
| Text in Textfeldern fehlt | Word | Textfelder mitlesen; Vollständigkeitsvergleich | ✅ |
| Automatische Nummerierung fehlt | Word | Nummernformat aus dem Dokument lesen; Vollständigkeitsvergleich | ✅ |
| Unbekannte Lücken im Word- oder PowerPoint-Leser | Word, PowerPoint | Vollständigkeitsvergleich gegen die Ansichts-PDF | ✅ |
| Zugeschnittenes Bild zeigt mehr als sichtbar | Word, PowerPoint | Warnvermerk, Seitenansicht | ✅ |
| Überschrift nicht erkannt | PDF | Regel im Skript, Claude prüft | ✅ / Claude |
| Bild an falscher Stelle | PDF | Vermerk, Claude verschiebt | ✅ / Claude |
| Fehlerhafte Texterkennung bei Scans | PDF | nur Claudes Blick auf die Seitenansicht | ⚠️ |
| Spalten durcheinander | PDF | Seitenansicht | ❓ |
| Foto seitlich | Bilder | beim Kopieren gerade richten | ✅ |
| Tippfehler im Original | alle | stehen lassen, `[unsicher:]` | ✅ |
| Anweisungen im Material („Ignoriere …") | alle | als Text behandeln, nie ausführen (Designprinzip 11) | ✅ Skill |

---

## 7. Die Sicherungen

`ablegen.py` legt **nur ab, was vollständig ist**. Es lehnt ab, mit Grund:

- Pflichtfelder fehlen: `fach`, `typ`, `thema`, `beschreibung`, `herkunft`
- `quelle` fehlt bei fremdem Material
- `url` oder `rechtsstand` fehlt bei Gesetzestexten
- ein Bild wurde nicht angesehen (Platzhalter steht noch)
- eine Scan-Seite wurde nicht abgeschrieben
- bei einer Bilddatei fehlen Beschreibung oder Abschrift
- dasselbe Original ist schon aufgenommen

Außerdem: Offene `[unsicher:]`- und `[Prüfen:]`-Stellen setzen den Status
automatisch auf `pruefen`. Nichts wird überschrieben; Originale werden nur
verschoben, nie gelöscht.

**Was die Sicherungen nicht leisten:** Sie prüfen, *dass* Claude
hingesehen hat, nicht, *ob richtig*. Fall 1 wäre durch alle Sicherungen
gekommen. Die fachliche Richtigkeit prüft am Ende nur die Lehrkraft –
deshalb landet alles Unsichere auf `pruefen`.

---

## 8. Tests

**Referenztests:** `tests/test_umwandlung.py` – nach jeder Änderung an der
Umwandlung laufen lassen. Die selbst gebauten Dateien unter
`tests/referenz/` halten fest:

| Datei | Prüft |
|---|---|
| `textfeld_und_nummerierung.docx` | Textfeld wird Kasten, automatische Nummern „1.–3." bleiben, kein Vollständigkeitsalarm |
| `zugeschnitten.docx` | Zuschnitt-Vermerk an einem Word-Bild |
| `ebenen.pptx` | Text über Bild erkannt, Zuschnitt erkannt, unauffälliges Bild ohne Vermerk |

Liegt echtes Material der Lehrkraft in `test_datein/` (nur lokal, nie im
Repo), prüft das Skript zusätzlich Fall 1 und die Word-Merkmale des
Vokabelblatts.

**Noch mit echtem Material zu testen:**

| Test | Warum |
|---|---|
| **Claude sieht PDF-Seiten in Cowork als Bild** | Grundlage der Seitenansicht – ohne sie fehlt die wichtigste Kontrolle |
| **Seitenansicht per LibreOffice in Cowork** | Hier nur über MS Word erprobt |
| Echter Scan und Handyfoto einer Seite | Häufigster Fall bei Lehrkräften |
| Mehrspaltiges PDF | Textreihenfolge |
| PowerPoint und Excel mit echtem Material | bisher nur mit selbst gebauten Dateien |
| `.pages`-Datei vom Mac der Lehrkraft | wahrscheinlich, ungetestet |
