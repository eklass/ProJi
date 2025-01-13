#!/bin/bash

# Skript zur Einrichtung und Ausführung des Tools
echo "Starte Setup für das Tool..."

# 1. Python-Version prüfen
if ! command -v python3 &> /dev/null
then
    echo "Python3 ist nicht installiert. Bitte installiere Python3, bevor du fortfährst."
	exit 1
fi

# 2. Virtuelle Umgebung erstellen
cd ..
echo "Erstelle virtuelle Umgebung..."
python3 -m venv venv

# 3. Virtuelle Umgebung aktivieren
echo "Aktiviere virtuelle Umgebung..."
source venv/bin/activate

# 4. Abhängigkeiten installieren
echo "Installiere Abhängigkeiten..."
pip install -r requirements.txt

# 5. Prüfen, ob Excel läuft
echo "Überprüfe, ob Excel-Prozesse aktiv sind..."

# Nach "Microsoft Excel" suchen und die PID ausgeben
PID=$(ps aux | grep -i "[M]icrosoft Excel" | awk '{print $2}')

if [ -n "$PID" ]; then
    echo "Bitte schließen Sie alle Excel-Fenster. Gefundener Prozess: PID $PID"
    while [ -n "$(ps aux | grep -i "[M]icrosoft Excel" | awk '{print $2}')" ]; do
        sleep 5
        PID=$(ps aux | grep -i "[M]icrosoft Excel" | awk '{print $2}')
        echo "Warte darauf, dass Excel geschlossen wird... Aktiver Prozess: PID $PID"
    done
fi

echo "Excel ist geschlossen. Setup wird fortgesetzt..."


# 6. Installiere xlwings Plugin für Excel
xlwings addin install

# 7. Installiere 1Password
brew install 1password-cli

# 8. Install terminal notifier
brew install terminal-notifier

# 9. Virtuelle Umgebung deaktivieren
echo "Deaktiviere virtuelle Umgebung..."
deactivate

echo "Setup abgeschlossen. 🍺"

