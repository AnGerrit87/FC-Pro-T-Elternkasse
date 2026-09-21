# FC Pro-T e.V. Elternkasse

## Google Apps Script mit clasp synchronisieren

Das Apps-Script-Backend liegt unter `apps-script/`.

### Einmalige Einrichtung

1. Node.js installieren.
2. clasp installieren: `npm install -g @google/clasp`
3. Bei Google anmelden: `clasp login`
4. `.clasp.json.example` nach `.clasp.json` kopieren.
5. Verbindung prüfen: `clasp status`
6. Backend zu Google Apps Script übertragen: `clasp push`

Die Datei `.clasp.json` sowie lokale OAuth-Daten werden durch `.gitignore` nicht committed.

### Wichtig

`clasp push` aktualisiert den Quellcode im Apps-Script-Projekt. Bei einem als Web-App bereitgestellten Apps Script kann anschließend eine neue Bereitstellung/Version erforderlich sein, damit die öffentliche Web-App die Änderung verwendet.
