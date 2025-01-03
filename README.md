# Zeitbuchungs-Tool "ProJi"
TLDR: Automatisierte Zeiterfassung für Jira und Projektron mit Selenium, Excel-Makros und Python, inklusive sicherer Passwortverwaltung via 1Password CLI

# "ProJi" Setup
Willkommen zum Zeitbuchungs-Tool! Dieses Tool ermöglicht es euch, über eine Excel-Datei Zeiten 
in JIRA und Projektron zu buchen.
Hier findet ihr eine Anleitung, wie ihr das Tool einrichtet und nutzt.

---

## Voraussetzungen

- **Python 3.x** installiert auf eurem System
    - Ihr könnt Python [hier herunterladen](https://www.python.org/downloads/).
- **Excel** mit aktiviertem Makrosupport

---

## Installation

1. **Dateien herunterladen**  
   Ladet die bereitgestellten Dateien (Excel-Datei, Python-Skripte und `installPythonLibraries.sh`) auf euren Computer.

2. **Virtuelle Umgebung einrichten**  
   Öffnet ein Terminal (macOS/Linux) im Ordner mit den heruntergeladenen Dateien und führt das Setup-Skript aus:

    - **macOS/Linux**:
      ```bash
      chmod +x installPythonLibraries.sh
      ./installPythonLibraries.sh
      ```

   Das Skript:
    - Erstellt eine virtuelle Python-Umgebung.
    - Installiert automatisch alle benötigten Abhängigkeiten um die Zeiten zu buchen.

---

## Nutzung

1. **Excel-Datei öffnen**  
   Öffnet die bereitgestellte Excel-Datei.

2. **Makros aktivieren**  
   Ein Popup sollte erscheinen, indem ihr gefragt werdet, ob ihr Makros aktivieren wollt.
   Klickt auf Ja, da die Makros notwendig sind um die Zeiten zu buchen.

3. **Zeiten buchen**  
   Füllt die entsprechenden Felder in der Excel-Datei aus und klickt auf die jeweiligen Buttons, um die Zeiten zu buchen.
   Eine Live-Demo gab es im Java Meeting (06.11.24). 
   Die Session wurde aufgezeichnet und kann bei der Bedienung des Excel Sheets helfen.

---

## Fehlerbehebung

### Python nicht installiert
Falls das Setup-Skript meldet, dass Python nicht installiert ist:
- Installiert Python von der offiziellen Website: [https://www.python.org/downloads/](https://www.python.org/downloads/).
- Wiederholt danach den Installationsschritt.

### Excel-Fehler: Makros deaktiviert
- Öffnet die Excel Datei erneut und klickt auf Makros aktivieren.

### Skript wird nicht ausgeführt
Falls das Skript beim Klick auf den Button nicht ausgeführt wird:
- Stellt sicher, dass das Setup-Skript ohne Fehler durchgelaufen ist.
- Überprüft, ob der Button korrekt mit dem Makro verknüpft ist.

---

## Kontakt

Bei Fragen oder Problemen meldet euch gerne bei mir.

Viel Erfolg beim Buchen! 🚀
