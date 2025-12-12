# RDD-AddIn Handbuch

**Room Design Document Add-In für Excel**  
*Version 0.10 – Für Adventure Game Studio (AGS) Spieleentwicklung*

---

## Inhaltsverzeichnis

1. [Einführung](#1-einführung)
2. [Installation](#2-installation)
3. [Die Benutzeroberfläche](#3-die-benutzeroberfläche)
   - 3.1 [Das Ribbon-Menü](#31-das-ribbon-menü)
   - 3.2 [Kontextmenüs](#32-kontextmenüs)
4. [Room Templates](#4-room-templates)
   - 4.1 [Struktur eines Room Sheets](#41-struktur-eines-room-sheets)
   - 4.2 [Räume erstellen und bearbeiten](#42-räume-erstellen-und-bearbeiten)
5. [Puzzle Dependency Chart (PDC)](#5-puzzle-dependency-chart-pdc)
   - 5.1 [Konzept und Methodik](#51-konzept-und-methodik)
   - 5.2 [PDC Workflow](#52-pdc-workflow)
   - 5.3 [Navigation im Chart](#53-navigation-im-chart)
6. [Listen-Synchronisation](#6-listen-synchronisation)
7. [Suche und Navigation](#7-suche-und-navigation)
8. [Export-Funktionen](#8-export-funktionen)
9. [Optionen und Einstellungen](#9-optionen-und-einstellungen)
10. [Tipps und Best Practices](#10-tipps-und-best-practices)

---

## 1. Einführung

Das **RDD-AddIn** (Room Design Document Add-In) ist eine umfassende Excel-Erweiterung für die Entwicklung von Adventure Games mit Adventure Game Studio (AGS). Es bietet eine strukturierte Methode zur Dokumentation von Räumen, Puzzles, Items, Actors und anderen Spielelementen.

Das Add-In basiert auf bewährten Game-Design-Methoden, insbesondere der **Puzzle Dependency Chart**-Methodik von Ron Gilbert, die bei der Entwicklung von Klassikern wie Monkey Island verwendet wurde.

### Hauptfunktionen

| Feature | Beschreibung |
|---------|--------------|
| Room Management | Erstellen, bearbeiten und verwalten von Room Design Documents |
| Puzzle Dependency Chart | Visualisierung von Puzzle-Abhängigkeiten nach Ron Gilbert |
| Dropdown-Synchronisation | Automatische Listen-Verwaltung aus Room-Daten |
| Kontextmenüs | Schnellzugriff auf häufige Funktionen |
| Find Usage | Suche nach Verwendungen von Items, Actors, Flags |
| Export | PDF- und CSV-Export der Dokumentation |

---

## 2. Installation

### Systemvoraussetzungen

- Microsoft Excel 2010 oder höher (Windows)
- Makros müssen aktiviert sein
- Scripting Runtime Library (scrrun.dll) – standardmäßig vorhanden

### Installationsschritte

**Schritt 1:** Kopieren Sie die Datei `RDD_AddIn.xlam` in den Excel Add-Ins Ordner:

```shell
%APPDATA%\Microsoft\AddIns\
```

**Schritt 2:** Öffnen Sie Excel und navigieren Sie zu:  
*Datei → Optionen → Add-Ins → Excel-Add-Ins verwalten → Los...*

**Schritt 3:** Aktivieren Sie das Kontrollkästchen neben "RDD_AddIn" und klicken Sie auf OK.

**Schritt 4:** Das neue Tab "RDD-AddIn" erscheint nun im Ribbon-Menü.

> 💡 **Info:** Beim ersten Start wird ein Arbeitsordner unter `%AppData%\BYTE RANGER\RDDAddIn` erstellt, der Log-Dateien und das Handbuch enthält, sowie ein temporärer Ordner unter `%Temp%\BYTE RANGER\RDDAddIn`.

---

## 3. Die Benutzeroberfläche

### 3.1 Das Ribbon-Menü

![Ribbon](images/Ribbon.png)
Nach der Installation erscheint ein neues Tab **RDD** im Excel-Ribbon mit folgenden Gruppen:

#### Gruppe: Räume

| Button | Funktion |
|--------|----------|
| **Add Room** | Erstellt ein neues Room Sheet basierend auf dem Template |
| **Bearbeiten** | Öffnet Dialog zur Bearbeitung von Room ID, Scene ID, Alias |
| **Löschen** | Löscht das aktuelle Room Sheet (mit Referenzprüfung) |
| **Sync Listen** | Synchronisiert alle Dropdown-Listen aus den Room-Daten |
| **Validieren** | Prüft Daten auf Duplikate, fehlende Referenzen, Zyklen |

#### Gruppe: Dependency Chart

| Button | Funktion |
|--------|----------|
| **Daten erstellen** | Extrahiert Puzzle-Daten und erstellt PDCData Sheet |
| **Chart generieren** | Erzeugt visuelles Puzzle Dependency Chart |
| **Chart aktualisieren** | Aktualisiert bestehendes Chart mit neuen Daten |

#### Gruppe: Export

| Button | Funktion |
|--------|----------|
| **PDF Export** | Exportiert Room Sheets als druckbares PDF |
| **CSV Export** | Exportiert PDC-Daten als CSV (nodes.csv, edges.csv) |

#### Gruppe: Info

| Button | Funktion |
|--------|----------|
| **Optionen** | Öffnet Einstellungen-Dialog |
| **Log** | Zeigt Log-Dateien an |
| **Handbuch** | Öffnet dieses Handbuch |
| **Version** | Zeigt About-Dialog mit Versionsinformationen |

### 3.2 Kontextmenüs

Das Add-In erweitert das Excel-Kontextmenü (Rechtsklick) mit kontextsensitiven Optionen. Je nach Position der aktiven Zelle werden unterschiedliche Menüoptionen angezeigt:

| Zelltyp | Menüoption 1 | Menüoption 2 |
|---------|--------------|--------------|
| **Room ID/Alias** | Neuen Raum anlegen | Zum Raum navigieren |
| **Puzzle ID** | Goto Node in Chart | Show Dependencies |
| **Item ID** | Find Usage | – |
| **Actor ID** | Find Usage | – |
| **Hotspot ID** | Find Usage | – |
| **Flag ID** | Find Usage | – |
| **Dependencies** | Goto Referenced | – |

---

## 4. Room Templates

### 4.1 Struktur eines Room Sheets

Jedes Room Sheet folgt einer standardisierten Struktur mit mehreren Abschnitten:

| Abschnitt | Zeilen | Inhalt |
|-----------|--------|--------|
| **ROOM HEADER** | 1 | Room ID, Scene ID, Room No, Room Alias |
| **CHECKLIST** | 3-12 | Status-Tracking für Assets (Backgrounds, Events, Speech, etc.) |
| **PICTURE AREA** | 3-12 | Platz für Screenshot oder Konzeptbild |
| **SCENE DESCRIPTION** | 15-23 | Narrative Beschreibung der Szene |
| **WHAT HAPPENS HERE?** | 24-38 | Story-Events und Gameplay-Ereignisse |
| **GENERAL SETTINGS** | 24-38 | Perspective, Parallax, Dimensionen, Viewport |
| **DOORS TO...** | 40-53 | Verbindungen zu anderen Räumen |
| **ACTORS** | 40-53 | Charaktere mit Conditions |
| **SOUNDS** | 55-68 | Sound Effects und Musik |
| **SPECIAL FX** | 55-68 | Animationen und Effekte |
| **PICKUPABLE OBJECTS** | 70-83 | Items zum Aufsammeln |
| **MULTI-STATE OBJECTS** | 70-83 | Objekte mit mehreren Zuständen |
| **TOUCHABLE OBJECTS** | 85-98 | Hotspots und interaktive Bereiche |
| **FLAGS / KNOWLEDGE** | 85-98 | Variablen und Wissens-Flags |
| **PUZZLES** | 100-115 | Vollständige Puzzle-Dokumentation |

#### PUZZLES Spalten

| Spalte | Beschreibung |
|--------|--------------|
| Puzzle ID | Eindeutige ID (z.B. P001, P002) |
| Title | Kurze Beschreibung des Puzzles |
| Target | Zielobjekt der Aktion |
| Action/Verb | Use, Talk, Give, Look, etc. |
| DependsOn | Vorausgesetzte Puzzles (kommagetrennt) |
| Requires | Benötigte Items/Flags |
| Grants | Gewährte Items/Flags nach Lösung |
| Difficulty | Schwierigkeitsgrad |
| Owner | Verantwortlicher Designer |
| Status | todo, in progress, done, n/a |
| Points | IQ-Punkte |
| Notes | Zusätzliche Notizen |

### 4.2 Räume erstellen und bearbeiten

#### Neuen Raum erstellen

1. Klicken Sie im Ribbon auf **„Add Room“**.
2. Optional: Geben Sie die Szenen-ID in das Dialogfeld ein (z. B. „Hindu-Tempel“).
3. Geben Sie einen Raumalias ein (z. B. „Eingang“).
4. Geben Sie eine AGS-Raumnummer ein (z. B. „1“).
5. Basierend auf der Vorlage wird ein neues Blatt erstellt.

> ⚠️ **Hinweis:** Raum-IDs müssen eindeutig sein und werden automatisch nach folgendem Schema „R###“ generiert.  
Dem Alias wird automatisch „r_“ vorangestellt.

#### Raum-Identität bearbeiten

1. Navigieren Sie zum gewünschten Room Sheet
2. Klicken Sie auf **"Bearbeiten"** im Ribbon
3. Ändern Sie Room ID, Scene ID oder Alias
4. Alle Referenzen werden automatisch aktualisiert

#### Raum löschen

1. Navigieren Sie zum zu löschenden Room Sheet
2. Klicken Sie auf **"Löschen"** im Ribbon
3. Bestätigen Sie die Löschung
4. Das System prüft vorher auf Referenzen in anderen Räumen

---

## 5. Puzzle Dependency Chart (PDC)

### 5.1 Konzept und Methodik

Das **Puzzle Dependency Chart** (PDC) ist eine visuelle Methode zur Darstellung von Puzzle-Abhängigkeiten in Adventure Games. Diese Technik wurde von Ron Gilbert entwickelt und bei klassischen LucasArts-Adventures wie "The Secret of Monkey Island" eingesetzt.

#### Node-Typen

| Node-Typ | ID-Präfix | Beschreibung | Farbe |
|----------|-----------|--------------|-------|
| Puzzle | P001, P002... | Ein lösbares Puzzle/Aufgabe | Blau |
| Item | i_key, i_map... | Ein Inventar-Gegenstand | Grün |
| Flag (Global) | g_doorOpen... | Globale Wissensvariable | Lila |
| Flag (Room) | r_visited... | Raumspezifische Variable | Orange |

#### Edge-Typen (Verbindungen)

| Edge-Typ | Spalte im Puzzle | Bedeutung |
|----------|------------------|-----------|
| depends | DependsOn | Puzzle X muss vor Puzzle Y gelöst werden |
| requires | Requires | Puzzle benötigt Item/Flag zur Lösung |
| grants | Grants | Puzzle gewährt Item/Flag nach Lösung |

### 5.2 PDC Workflow

```txt
┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐
│ Room Sheets │───>│ Validierung │───>│ PDC Daten   │───>│   Chart     │───>│ Navigieren  │
│ mit Puzzles │    │durchführen  │    │  erstellen  │    │ generieren  │    │& Analysieren│
│   befüllen  │    │             │    │             │    │             │    │             │
└─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘
```

#### Schritt 1: Puzzles dokumentieren

- Puzzle ID eingeben
- DependsOn definieren
- Requires festlegen
- Grants zuweisen

#### Schritt 2: Validierung

- Duplikate prüfen
- IDs validieren
- Referenzen checken
- Zyklen erkennen

#### Schritt 3: Daten erstellen

- Nodes extrahieren
- Edges erstellen
- Types zuweisen
- Sheet "PDCData" wird erstellt

#### Schritt 4: Chart generieren

- Shapes erzeugen
- Connectors ziehen
- Layout anwenden
- Sheet "Chart" wird erstellt

#### Schritt 5: Navigation

- Ctrl+Click auf Node → Zur Quelle springen
- Dependencies analysieren

### 5.3 Navigation im Chart

Das generierte Chart ist vollständig interaktiv:

**Ctrl+Klick auf einen Node:**  
Springt direkt zur Puzzle-Definition im entsprechenden Room Sheet. Dies verwendet die Windows API (GetAsyncKeyState) zur Erkennung der Ctrl-Taste beim Klick.

**Kontextmenü auf PDCData:**  
Bei Rechtsklick auf eine Puzzle-Zelle im PDCData-Sheet erscheinen zusätzliche Optionen:

- "Goto Node in Chart" – Zum Node im Chart navigieren
- "Show Dependencies" – Alle Abhängigkeiten anzeigen

---

## 6. Listen-Synchronisation

Das Add-In verwaltet automatisch die Dropdown-Listen im Dispatcher-Sheet. Diese Listen werden für Validierung und Auto-Complete in den Room Sheets verwendet.

### Verwaltete Listen

| Liste | Quelle | Verwendung |
|-------|--------|------------|
| Room ID | Alle Room Sheets | DOORS TO... Navigation |
| Room Alias | Alle Room Sheets | DOORS TO... Navigation |
| Scene ID | Alle Room Sheets | Referenzierung |
| Actor ID | ACTORS-Bereiche | Puzzle Owner/Target |
| Actor Name | ACTORS-Bereiche | Anzeige |
| Item ID | PICKUPABLE OBJECTS | Requires/Grants |
| Item Name | PICKUPABLE OBJECTS | Anzeige |
| Flag ID | FLAGS-Bereiche | Requires/Grants |
| Hotspot ID | TOUCHABLE OBJECTS | Puzzle Target |
| Puzzle ID | PUZZLES-Bereiche | DependsOn |

### Automatische Synchronisation

Bei aktivierter Option "Auto Sync Lists" werden die Listen automatisch aktualisiert wenn:

- Ein Room Sheet geändert wird
- Ein neuer Raum erstellt wird
- Ein Raum gelöscht wird

### Manuelle Synchronisation

Der Button **"Synchronize Lists"** im Ribbon erzwingt eine vollständige Synchronisation.

Der Button zeigt zwei Zustände:

- ![🟢 **Grün:**](images/SyncGreen.png)  Listen sind synchron
- ![🟠 **Orange:**](images/SyncOrange.png) Änderungen erkannt, Sync empfohlen

---

## 7. Suche und Navigation

Die "Find Usage"-Funktion ermöglicht das Auffinden aller Verwendungen von Items, Actors, Hotspots und Flags über alle Room Sheets hinweg.

### Find Usage aufrufen

1. Positionieren Sie den Cursor auf einer ID-Zelle (z.B. Item ID)
2. Rechtsklick → "Find Usage" wählen
3. Das Suchergebnis-Fenster öffnet sich
4. Doppelklick auf ein Ergebnis navigiert zur entsprechenden Zelle

### Durchsuchte Bereiche

| Element | Durchsuchte Spalten |
|---------|---------------------|
| Items | Puzzles_Requires, Puzzles_Grants, PickupableObjects_ItemID |
| Actors | Actors_Condition, Puzzles_Owner, Puzzles_Target |
| Hotspots | TouchableObjects_HotspotID, Puzzles_Target |
| Flags | Flags_FlagID, Puzzles_Requires, Puzzles_Grants |

> 💡 **Tipp:** Die Suche unterstützt kommaseparierte Werte in Zellen. Wenn eine Zelle `i_key, i_map` enthält, wird bei Suche nach `i_key` diese Zelle als Treffer angezeigt.

---

## 8. Export-Funktionen

### PDF Export

Der PDF-Export erstellt ein druckbares Dokument mit allen Room Sheets:

- Jedes Room Sheet wird als eigene Seite exportiert
- Formatierung und Bilder werden beibehalten
- Optimiert für A4-Querformat

**Aufruf:** Ribbon → Export → "PDF Export"  
**Speicherort:** Dialog zur Auswahl des Zielordners

### CSV Export

Der CSV-Export erstellt separate Dateien für die PDC-Daten:

- `nodes.csv` – Alle Puzzle-Nodes
- `edges.csv` – Alle Abhängigkeiten

Diese Dateien können in anderen Tools (z.B. Graphviz, yEd) zur weiteren Visualisierung verwendet werden.

---

## 9. Optionen und Einstellungen

Das Optionen-Fenster (Ribbon → Info → "Optionen") bietet zwei Bereiche:

### Allgemeine Einstellungen (Registry)

| Einstellung | Beschreibung | Standard |
|-------------|--------------|----------|
| Manual Path | Pfad zum Handbuch-Verzeichnis | `%AppData%\BYTE RANGER\RDDAddIn\` |
| Log Retention Days | Tage bis alte Logs gelöscht werden | 30 |

### Arbeitsmappe-Einstellungen (Document Properties)

| Einstellung | Beschreibung | Standard |
|-------------|--------------|----------|
| Default Game Width | Standard-Spielbreite in Pixeln | 320 |
| Default Game Height | Standard-Spielhöhe in Pixeln | 200 |
| Default BG Width | Standard-Hintergrundbreite | 320 |
| Default BG Height | Standard-Hintergrundhöhe | 200 |
| Default UI Height | Standard-UI-Höhe | 40 |
| Default Perspective | Standard-Perspektive | (leer) |
| Default Parallax | Standard-Parallax-Modus | None |
| Default Scene Mode | Standard-Szenen-Modus | (leer) |
| Auto Sync Lists | Automatische Listen-Synchronisation | True |
| Show Validation Warnings | Validierungswarnungen anzeigen | True |

---

## 10. Tipps und Best Practices

### Namenskonventionen

| Element       | Konvention                  | Beispiel             | Behandelt |
|---------------|-----------------------------|----------------------|-----------------------|
| Room ID       | R + dreistellige Nummer     | R001, R002, R100     | ✅ Automatisch        |
| Room Alias    | r_ + beschreibender Name    | r_entrance, r_cellar | ✅ Automatisch        |
| Puzzle ID     | P + dreistellige Nummer     | P001, P002           | ✅ Automatisch        |
| Item ID       | i_ + Name                   | i_key, i_goldcoin    | ✅ Automatisch        |
| Flag (Global) | g_ + Name                   | g_doorUnlocked       | ✅ Automatisch        |
| Flag (Room)   | r_ + Name                   | r_visited            | ✅ Automatisch        |
| Actor ID      | c + Name (Character)        | cEgo, cBartender     | ✅ Automatisch        |
| Hotspot ID    | h + Name                    | hDoor, hWindow       | ✅ Automatisch        |
| State Object  | o + Name                    | oDoor, oLever        | ✅ Automatisch        |

> 👉 Hinweis: "✅ Automatisch" bedeutet, dass das Add-In beim Erstellen des Elements die ID oder den Namen direkt nach der Konvention vergibt. Bei "⚠️ Manuell" muss der Benutzer selbst darauf achten, die richtige Schreibweise einzuhalten.

### Workflow-Empfehlungen

- **Regelmäßig validieren:** Führen Sie nach größeren Änderungen immer eine Validierung durch.
- **Listen synchron halten:** Bei deaktiviertem Auto-Sync regelmäßig manuell synchronisieren.
- **Backups erstellen:** Vor großen Änderungen eine Kopie der Arbeitsmappe anlegen.
- **PDC iterativ aufbauen:** Beginnen Sie mit den Haupt-Puzzles und verfeinern Sie später.
- **Konsistente IDs:** Verwenden Sie durchgängig die gleichen Namenskonventionen.

### Fehlerbehebung

| Problem | Lösung |
|---------|--------|
| Ribbon erscheint nicht | Add-In erneut aktivieren unter Excel-Optionen → Add-Ins |
| Buttons ausgegraut | Stellen Sie sicher, dass eine RDD-Arbeitsmappe geöffnet ist |
| Validierungsfehler | Prüfen Sie die Log-Datei unter Info → Log |
| Chart nicht aktualisiert | "Chart aktualisieren" nach Datenänderungen ausführen |
| Listen nicht synchron | Manuell "Sync Listen" ausführen |
| Kontextmenü erscheint nicht | Cursor auf gültige ID-Zelle positionieren |

---

*RDD-AddIn – Room Design Document Add-In für Adventure Game Studio*  
*Dokumentation Version 0.10*
