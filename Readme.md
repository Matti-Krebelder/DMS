# DMS - Device Management System

Ein umfassendes Geräteverwaltungssystem für Schulen und Unternehmen, entwickelt mit Flask und SQLite.

## 🚀 Installation und Setup

### Voraussetzungen

- Python 3.7 oder höher
- pip (Python Package Manager)

### Installation

1. **Repository klonen:**
   ```bash
   git clone https://github.com/Matti-Krebelder/DMS.git
   cd DMS
   ```

2. **Virtuelle Umgebung erstellen (empfohlen):**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # Auf Windows: venv\Scripts\activate
   ```

3. **Abhängigkeiten installieren:**
   ```bash
   pip install flask flask-cors requests qrcode pillow python-docx reportlab
   ```

4. **Anwendung starten:**
   ```bash
   python3 app.py
   ```

5. **Im Browser öffnen:**
   Öffnen Sie http://localhost:5000 in Ihrem Webbrowser.

## 📋 Funktionen

### 🔐 Benutzerverwaltung
- Sichere Anmeldung mit Benutzer-ID
- Sitzungsverwaltung
- Rollenbasierte Zugriffsrechte

### 🏭 Lagerverwaltung
- Mehrere Lager erstellen und verwalten
- Persönliche und schulische Lager-Systeme
- Zugriffsrechte für verschiedene Benutzer

### 📦 Geräteverwaltung
- Geräte hinzufügen, bearbeiten und löschen
- Automatische Barcode-Generierung
- Detaillierte Geräte-Informationen:
  - Name, Barcode, Lagerplatz
  - Beschreibung, Seriennummer, Modell
  - Instrumentenart, Inventarnummer
  - Kaufdatum, Preis

### 🔄 Ausleihsystem
- Geräte ausleihen und zurückgeben
- QR-Code basierte Rückgabe
- Ausleihhistorie und -details
- PDF-Generierung für Ausleihübersichten
- E-Mail und Klassen-Tracking

### 📊 Inventarverwaltung
- Umfassende Such- und Filterfunktionen
- Gruppierung nach Modell, Instrument, Status
- Export-Funktionen:
  - CSV-Export
  - Word-Dokument Export
  - PDF-Etiketten mit QR-Codes

### 🏷️ Label-System
- Benutzerdefinierte Etiketten-Layouts
- QR-Code Integration
- Druckoptimierte PDF-Generierung

### 🔍 Geräte-Scanner
- QR-Code basierte Gerätesuche
- Sofortige Geräte- und Ausleihinformationen
- Integration mit dem Ausleihsystem

### 📈 Dashboard
- Übersicht über alle Lager
- Update-Benachrichtigungen
- Schnellzugriff auf wichtige Funktionen

### 🔄 Automatische Updates
- Versionsprüfung
- One-Click Update-Funktionalität
- Automatische Datei-Aktualisierung

## 🗂️ Projektstruktur

```
DMS/
├── app.py                 # Hauptanwendung
├── users.db              # Benutzer- und Lager-Datenbank
├── templates/            # HTML-Templates
│   ├── login.html
│   ├── dashboard.html
│   ├── warehouse.html
│   ├── devices.html
│   ├── add_device.html
│   ├── edit_device.html
│   ├── borrow.html
│   ├── borrow_success.html
│   ├── return.html
│   ├── inventory.html
│   ├── manage_lager.html
│   ├── edit_lager.html
│   ├── create_lager.html
│   ├── export.html
│   ├── label_selection.html
│   ├── label_layout.html
│   └── info.html
├── backups/              # Automatische Datenbank-Backups
├── images/               # Gerätebilder (optional)
└── *.db                  # Lager-spezifische Datenbanken
```

## 🛠️ Technische Details

### Datenbanken
- **users.db**: Globale Benutzer- und Lager-Informationen
- **{lager_id}.db**: Lager-spezifische Geräte- und Ausleihdaten

### Abhängigkeiten
- **Flask**: Web-Framework
- **Flask-CORS**: Cross-Origin Resource Sharing
- **Requests**: HTTP-Anfragen für Updates
- **QRCode**: QR-Code Generierung
- **Pillow**: Bildverarbeitung
- **python-docx**: Word-Dokument Erstellung
- **ReportLab**: PDF-Generierung

### Sicherheit
- Sitzungsbasierte Authentifizierung
- CSRF-Schutz durch Flask-WTF
- Sichere Passwort-Verwaltung (empfohlen für Produktion)

## 📖 Verwendung

### Erste Schritte
1. Nach der Installation die Anwendung starten
2. Mit einer bestehenden Benutzer-ID anmelden (Standard: CKS.EXampleid)
3. Ein neues Lager erstellen
4. Geräte hinzufügen und verwalten

### Tägliche Nutzung
- **Geräte hinzufügen**: Über "Gerät hinzufügen" neue Geräte registrieren
- **Ausleihen**: Geräte über das Ausleihsystem verleihen
- **Rückgaben**: QR-Codes für schnelle Rückgaben verwenden
- **Inventur**: Über "Inventar" den Bestand überprüfen
- **Export**: Daten in verschiedenen Formaten exportieren

## 🔧 Konfiguration

### Umgebungsvariablen
```bash
export FLASK_ENV=development  # Für Entwicklung
export FLASK_DEBUG=1          # Debug-Modus aktivieren
```

### Datenbank-Backups
Die Anwendung erstellt automatisch Backups bei kritischen Operationen:
- Geräte hinzufügen/bearbeiten/löschen
- Ausleihen und Rückgaben

Backups werden im `backups/` Ordner gespeichert.

## 🤝 Beitragen

1. Fork das Repository
2. Erstelle einen Feature-Branch (`git checkout -b feature/AmazingFeature`)
3. Commit deine Änderungen (`git commit -m 'Add some AmazingFeature'`)
4. Push zum Branch (`git push origin feature/AmazingFeature`)
5. Öffne einen Pull Request

## 📝 Lizenz

Dieses Projekt ist unter der MIT-Lizenz lizenziert - siehe die [LICENSE](LICENSE) Datei für Details.

## 📞 Support

Bei Fragen oder Problemen:
- Öffne ein Issue auf GitHub
- Kontaktiere den Entwickler

## 🔄 Updates

Die Anwendung prüft automatisch auf neue Versionen und bietet One-Click Updates an. Updates beinhalten:
- Neue Funktionen
- Bugfixes
- Sicherheitsverbesserungen
- Template-Updates

---
