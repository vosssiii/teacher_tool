---
name: laufzeit-test
description: Prüft, ob der Unterrichtsassistent auf diesem Computer wirklich arbeiten kann – Word, PowerPoint, PDF, Bilder und Internetzugang. Nutze das, wenn jemand wissen will, ob das Plugin läuft, ob alles bereit ist, ob der Rechner das schafft, ob Arbeitsblätter und Klassenarbeiten hier erstellt werden können, oder wenn nach einem Selbsttest, Systemcheck, Funktionstest oder Laufzeit-Test gefragt wird. Auch bei Sätzen wie "geht das hier überhaupt", "funktioniert das bei mir", "teste mal die Umgebung", "kann ich loslegen".
---

# Laufzeit-Test (Phase 1a)

Dieser Skill beantwortet die neun Testpunkte T1–T9 aus dem Konzeptpapier
(Kap. 13). Das Papier liegt **nicht** im Plugin – alles, was du brauchst,
steht hier. Such nicht danach.
Er ist **Wegwerf-Werkzeug**: Sobald Phase 1a abgeschlossen ist, wird er
durch die echten Skills (`einrichten`, `aufnehmen`, `material`, …) ersetzt.

## Der Grundgedanke

Vier Testpunkte lassen sich messen, fünf nur beobachten:

| Beantwortet das Skript | Beantwortest **du** durch Beobachtung |
| --- | --- |
| T1 Bibliotheken | T4 Befehlsnamen und Skriptaufruf |
| T2 Word → PDF und Ersatzweg | T5 Modellwahl |
| T3 Word/PowerPoint aus Vorlage | T7 Ordner-Anweisung |
| T6 Webzugriff im Skript | T8 freie Formulierungen |
| | T9 Dateien pro Durchlauf |

Beides landet in **einem** Bericht. Solange bei den Beobachtungen „offen“
steht, ist Phase 1a nicht abgeschlossen.

**Erfinde keine Beobachtung.** Was du nicht sicher feststellen konntest,
bleibt „konnte ich nicht feststellen“ – mit Begründung. Ein ehrliches
Fragezeichen ist für Phase 1a mehr wert als eine geratene Antwort, denn
auf diesen Bericht stützt sich die Entscheidung, ob weitergebaut wird.

## Schritt 1 – Beobachte, bevor du irgendetwas tust

Halte **jetzt** fest, solange es frisch ist, und merke es dir für Schritt 4:

- **T4:** Womit wurde dieser Skill ausgelöst? Notiere den Text der letzten
  Nachricht **wörtlich**. War es ein Befehl mit Schrägstrich? Wie lautete
  er genau, mit oder ohne Plugin-Präfix? Oder war es ein freier Satz?
- **T8:** Wenn es ein freier Satz ohne Befehl war, ist T8 damit schon zur
  Hälfte beantwortet – schreibe den Satz mit.
- **T5:** Mit welchem Modell läufst du gerade? Wenn du es nicht sicher
  weißt, schreibe das genau so hin.

## Schritt 2 – Skript ausführen, in **beiden** Umgebungen

Cowork kann Code an zwei verschiedenen Orten ausführen, und sie verhalten
sich grundverschieden:

| Ort | Dateien | Erkennbar an |
| --- | --- | --- |
| **lokale Geräte-Sandbox** (`device_bash`) | Skripte arbeiten direkt im Ordner der Lehrkraft | kein `/mnt/user-data` |
| **Cloud-Umgebung** | Dateien müssen einzeln herein (`device_stage_files`) und heraus (`device_commit_files`) | `/mnt/user-data` vorhanden |

**Führe den Test nach Möglichkeit in beiden aus** und gib bei `--umgebung`
an, welcher es war. Ein Bericht aus der Cloud sagt nichts über die lokale
Sandbox: andere Architektur (ARM statt x86), möglicherweise anderes
Abbild, andere Programme. Das Skript schreibt seinen vermuteten
Ausführungsort selbst in den Kopf des Berichts – prüfe, ob er zu deiner
Erwartung passt.

Scheitert `device_bash` mit „Workspace unavailable“, halte das fest: Auf
dem Rechner der Lehrkraft fehlt dann die Virtualisierung. Auf Apple
Silicon ist sie immer vorhanden, unter Windows muss sie im BIOS
eingeschaltet sein.



Das Skript liegt im Plugin unter `scripts/laufzeit_test.py`.

Führe es aus und **schreibe mit, welcher Weg funktioniert hat** – das ist
die zweite Hälfte von T4:

```bash
python "${CLAUDE_PLUGIN_ROOT}/scripts/laufzeit_test.py" --umgebung "Claude Cowork"
```

Wenn `${CLAUDE_PLUGIN_ROOT}` nicht aufgelöst wird, probiere der Reihe nach
und notiere, was geklappt hat:

1. `python3` statt `python`
2. den absoluten Pfad zum Plugin-Ordner
3. Plugin-Ordner suchen (z. B. unter `~/.claude/plugins/`)

**Wenn gar kein Skript ausgeführt werden kann,** ist das der wichtigste
Befund des ganzen Tests: Designprinzip 3 („Skript vor Sprachmodell“) trägt
dann nicht, und die Architektur muss neu bewertet werden, bevor
weitergebaut wird. Halte in dem Fall genau fest, woran es scheiterte, und
gehe trotzdem zu Schritt 4 – der Bericht wird dann von Hand angelegt.

Das Skript schreibt nach `./laufzeit-test/`:
- `bericht.md` – der lesbare Bericht
- `roh.json` – Rohdaten zum Vergleich späterer Läufe

Es läuft immer durch und endet immer mit Erfolg, auch wenn einzelne Proben
scheitern. Ein Fehlschlag ist ein Ergebnis, kein Absturz.

**Fehlen unverzichtbare Bibliotheken,** frag die Lehrkraft kurz, ob du sie
nachinstallieren darfst, und rufe das Skript dann noch einmal so auf:

```bash
python "${CLAUDE_PLUGIN_ROOT}/scripts/laufzeit_test.py" --umgebung "Claude Cowork" --install
```

Das beantwortet die zweite Hälfte von T1 („bzw. lassen sich installieren“).

**Danach ein drittes Mal ohne `--install` laufen lassen.** Das ist kein
überflüssiger Lauf: Frisch installierte Pakete sind teils erst in einem
neuen Prozess voll nutzbar. `docx2pdf` etwa meldet im Installationslauf
noch einen Fehler und erzeugt im nächsten Lauf ein einwandfreies PDF – ein
Lauf direkt nach `--install` bewertet T2 also zu streng. Es zählt der
letzte Lauf.

## Schritt 3 – Bericht lesen

Lies `laufzeit-test/bericht.md`. Die Ampel oben zeigt T1, T2, T3 und T6.
Fasse das Ergebnis in **einfacher Sprache** zusammen, ohne Technikbegriffe
(Designprinzip 7). Nicht „python-docx fehlt“, sondern „Word-Dateien kann
ich hier im Moment nicht erzeugen“.

## Schritt 4 – Die fünf Beobachtungen feststellen

Nur diese Punkte fehlen jetzt noch. Arbeite sie der Reihe nach ab.

### T4 – Befehle und Skriptaufruf
Aus Schritt 1 und 2 weißt du das meiste. Ergänze:
- Musste die Ausführung des Skripts bestätigt werden, oder lief es sofort?
- Falls es ein Befehlsmenü gibt: Wie heißt der Skill dort?

### T5 – Modellwahl
- Welches Modell läuft gerade?
- Sieh im Frontmatter dieser Datei nach: Es enthält **kein** Feld `model:`.
  Prüfe, ob die Laufzeitumgebung ein solches Feld überhaupt kennt. Wenn du
  es nicht sicher feststellen kannst, schreibe genau das.
- Gibt es in Cowork eine sichtbare Modellauswahl? Wenn ja, wo?

### T6b – Webzugriff von Claude selbst
Das Skript hat nur geprüft, ob **es** ins Netz kommt. Prüfe getrennt, ob
**du** ein Werkzeug für Websuche oder Seitenabruf hast. Wenn ja, ruf damit
`https://www.gesetze-im-internet.de/bgb/__622.html` ab und notiere das
Ergebnis. Wenn nein, halte fest, dass die Rechtsstand-Prüfung aus Kap. 8.2
in Version 1 nicht automatisch geht.

### T7 – Ordner-Anweisung
Kann das Plugin die Ordner-Anweisung selbst setzen? Prüfe, ob du eine
Anweisungsdatei im Arbeitsordner anlegen kannst (etwa `CLAUDE.md`), und ob
Cowork sie liest. **Lege nichts ungefragt an** – frag vorher, und halte
sonst fest, wie viele Schritte die Lehrkraft von Hand tun müsste.

### T8 – Freie Formulierungen
Bitte die Lehrkraft, in **einer neuen Unterhaltung** nacheinander diese
Sätze zu schreiben, ohne Schrägstrich-Befehl:

1. „Läuft das Plugin hier eigentlich?“
2. „Kannst du prüfen, ob mein Rechner das schafft?“
3. „Ich will wissen, ob ich mit dem Unterrichtsassistenten loslegen kann.“
4. „Mach mal einen Systemcheck.“

Notiere für jeden Satz, ob dieser Skill ohne Befehl angesprungen ist.

> Das ist der Test, der später über `/material` und `/aufnehmen`
> entscheidet: Greift ein Skill nur auf Befehl, muss die Lehrkraft Befehle
> auswendig lernen – gegen Designprinzip 9.

### T9 – Umfang pro Durchlauf
- Wie lange hat dieser Lauf gedauert?
- Wie viel vom Kontingent hat er verbraucht (Anzeige in Cowork)?
- Leite daraus eine Empfehlung ab, wie viele Dateien ein
  `/aufnehmen`-Durchlauf verarbeiten sollte. Startwert im Konzept: 10.

## Schritt 5 – Beobachtungen in den Bericht eintragen

Öffne `laufzeit-test/bericht.md` und ersetze im Abschnitt
„Beobachtungen aus der Sitzung“ jedes `  - offen` durch die Antwort.
Trage bei „Ergebnis auf einen Blick“ außerdem für T4, T5, T7, T8 und T9
das ⬜ gegen ✅, ⚠️ oder ❌ aus.

Fülle zuletzt das **Fazit** aus. Es beantwortet genau eine Frage aus
Kap. 11: *Kann so weitergebaut werden, oder müssen Architektur und
Designprinzip 3 neu bewertet werden?*

Empfehle „neu bewerten“, wenn eines davon zutrifft:
- Skripte lassen sich gar nicht ausführen,
- weder Word noch PowerPoint lassen sich erzeugen,
- kein Weg führt zu einem PDF,
- der Skill greift bei keiner freien Formulierung.

## Schritt 6 – Zusammenfassen

Berichte der Lehrkraft in wenigen Sätzen, einfache Sprache: was geht, was
nicht, und was als Nächstes zu tun ist. Nenne den Pfad zum Bericht.

Sag ausdrücklich dazu, **was dieser Test nicht prüft**: die Qualität der
erzeugten Materialien. Er prüft nur, ob die Technik hier trägt.
