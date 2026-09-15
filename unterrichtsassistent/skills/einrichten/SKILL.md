---
name: einrichten
description: Richtet den Arbeitsordner für Unterrichtsmaterialien ein oder bringt ihn auf den neuesten Stand. Nutze das, wenn jemand mit dem Unterrichtsassistenten anfangen will, den Ordner vorbereiten, die Schule oder Klassen hinterlegen, eine Word-Vorlage einrichten, den Notenschlüssel oder die Operatorenliste hinterlegen möchte – oder nach einem Update des Plugins. Auch bei Sätzen wie "ich will loslegen", "richte mir das ein", "hier ist meine Vorlage", "neues Schuljahr", "wo trage ich meine Klassen ein".
---

# Einrichten

Bereitet den Arbeitsordner der Lehrkraft vor (Konzept Kap. 6.1). Der Befehl
ist **wiederholbar** – zum neuen Schuljahr, nach einem Plugin-Update oder
wenn Unterlagen nachgereicht werden. Beim zweiten Mal wird nur aktualisiert,
was sich geändert hat.

## Grundregeln

**Überschreibe nichts von ihr.** Nur `_system/` gehört dem Plugin. Alles
andere – `_kontext/`, `wissensbasis/`, `entwuerfe/`, `originale/` – wird
angelegt, wenn es fehlt, und sonst in Ruhe gelassen.

**Frag wenig** (Designprinzip 5). Das Gespräch in Schritt 3 ist kurz. Was
sie nicht beantwortet, bleibt leer – nicht geraten.

**Sprich einfach** (Designprinzip 7). Keine Dateipfade, keine Fachbegriffe,
wenn es auch ohne geht. Nicht „`_kontext/notenschluessel.md` fehlt", sondern
„Für Klassenarbeiten brauche ich noch deinen Notenschlüssel".

## Schritt 1 – Arbeitsordner bestimmen

Der Arbeitsordner ist der Ordner, den sie **angehängt** hat. Prüfe, ob einer
angehängt ist.

**Ist keiner angehängt, brich hier ab** und sag ihr in einfachen Worten:
Sie muss den Ordner beim Start einer Aufgabe verbinden, sonst arbeitest du
ins Leere. Das ist die häufigste Fehlerquelle überhaupt.

Ist der Ordner leer und heißt noch nicht nach ihrem Unterricht, frag kurz,
ob er der richtige ist.

## Schritt 2 – Skript ausführen

Es erledigt alles, was immer gleich läuft: Ordner anlegen, Skripte und
Vorlagen bereitstellen, Inhaltsverzeichnis und Prüfwerte anlegen,
Bereitschaft je Befehl feststellen.

```bash
python "<Plugin>/scripts/einrichten.py" --ordner "<Arbeitsordner>"
```

`${CLAUDE_PLUGIN_ROOT}` ist in Cowork **leer** – verlass dich nicht darauf.
Such den Plugin-Ordner (unter `~/.claude/plugins/`) und nimm den vollen
Pfad. Klappt `python` nicht, nimm `python3`.

Läuft das Skript in der lokalen Geräte-Sandbox und liegt das Plugin dort
nicht, hol die Skriptdatei zuerst dorthin. Ab dem zweiten Aufruf liegt sie
ohnehin unter `_system/skripte/` im Arbeitsordner – nimm dann die.

Mit `--json` bekommst du dasselbe Ergebnis maschinenlesbar.

## Schritt 3 – Angaben zur Schule

Nur nötig, wenn `_kontext/schule.md` **fehlt**. Ist sie da, überspring
diesen Schritt und erwähne ihn nur, falls etwas offensichtlich veraltet ist
(etwa ein vergangenes Schuljahr).

Führe ein **kurzes** Gespräch. Frag in einer einzigen Nachricht nach:

- Name und Schulform der Schule
- Bildungsgang
- Schuljahr
- Welche Klassen in welchem Fach
- **Präsentationen als PowerPoint oder als HTML-Folien?** Erklär den
  Unterschied in einem Satz: PowerPoint zum Weiterbearbeiten, HTML zum
  Vorführen im Browser.
- Bei Englisch: Erwartungshorizont auf Deutsch oder auf Englisch?

Schreib daraus `_kontext/schule.md` nach dem Aufbau in
`_system/vorlagen/schule_vorlage.md`. Was sie nicht beantwortet hat, bleibt
leer stehen.

## Schritt 4 – Mitgebrachte Unterlagen

Liegen Dateien in `eingang/`, die in den Kontext gehören – Notenschlüssel,
Operatorenliste, Vorgabe zur Punkteverteilung, Word- oder
PowerPoint-Vorlage –, verarbeite sie jetzt:

| Datei | Ziel |
| --- | --- |
| Notenschlüssel | `_kontext/notenschluessel.md` |
| Operatorenliste | `_kontext/operatoren.md` |
| Vorgabe Punkteverteilung / Anforderungsbereiche | `_kontext/anforderungsbereiche.md` |
| Word-Vorlage (.docx/.dotx) | `_kontext/vorlage.docx` |
| PowerPoint-Vorlage (.pptx/.potx) | `_kontext/vorlage.pptx` |

Wandle Textdokumente in Markdown um und leg das Original nach `originale/`.
Vorlagen bleiben Vorlagen – die werden kopiert, nicht umgewandelt.

**Unterrichtsmaterialien gehören nicht hierher.** Arbeitsblätter, Texte,
alte Klassenarbeiten lässt du liegen und verweist am Ende auf `/aufnehmen`.

## Schritt 5 – Vorlage prüfen

Nur wenn eine Vorlage in `_kontext/` liegt:

```bash
python "<Skripte>/vorlage_pruefen.py" --datei "<Arbeitsordner>/_kontext/vorlage.docx"
```

Das Skript sucht je Zweck (Titel, Aufgabe, Fließtext, Quelle …) die erste
brauchbare Formatvorlage und fällt dabei auf die gewöhnlichen Word-Namen
zurück. Eine normale Schulvorlage besteht das fast immer.

- **Brauchbar** → sag ihr, dass ihre Vorlage benutzt wird.
- **Nicht brauchbar** → das Plugin nimmt seine Standardvorlage. Sag ihr in
  einem Satz, was fehlte, und dass ihre Datei nicht verändert wurde.

Gleiches für `vorlage.pptx`, falls vorhanden.

## Schritt 6 – Ordner-Anweisung

Gib den Text aus `_system/vorlagen/ordner_anweisung.md` aus – alles
unterhalb der Trennlinie – und erklär in einfachen Schritten, wo sie ihn
einträgt: in die Anweisung ihres Projekts in den Einstellungen.

**Sag ehrlich dazu, dass du das nicht für sie tun kannst.** Es gibt keinen
Weg, eine ordnerbezogene Anweisung automatisch zu setzen. Leg dafür auch
keine Datei im Arbeitsordner an – das wirkt nicht und stiftet nur
Verwirrung.

Hat sie das schon einmal gemacht, frag nur kurz nach, ob der Text noch
steht, statt ihn erneut auszugeben.

## Schritt 7 – Bericht

Sag ihr in wenigen Sätzen, was jetzt geht und was nicht. Halte dich an die
Bereitschaft aus dem Skript und nenn je Befehl genau das, was fehlt:

> Die Einrichtung ist fertig.
>
> ✅ Du kannst Material aufnehmen und Arbeitsblätter erstellen.
> ⚠️ Für Klassenarbeiten fehlt noch deine Operatorenliste und die Vorgabe,
> wie die Punkte auf die Anforderungsbereiche verteilt werden sollen. Leg
> beides in den Ordner `eingang` und ruf mich noch einmal auf.

Fehlt etwas, ist **nur der betroffene Befehl** eingeschränkt – sag das
ausdrücklich, damit sie nicht denkt, es gehe gar nichts.

Liegen noch Unterrichtsmaterialien in `eingang/`, verweise zum Schluss in
einem Satz auf `/unterrichtsassistent:aufnehmen`.
