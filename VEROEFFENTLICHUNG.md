# Automatische Veröffentlichung für Bibel-Worte

Das öffentliche Repository übernimmt fällige Inhalte aus dem privaten Repository
`AndreasHardt/bibel-worte-eingang`.

## Eingangsdateien

Pro Thema liegen im Ordner `inhalte` genau vier Dateien:

- `YYMMDD Thema.docx`
- `YYMMDD Thema DEU.jpg`
- `YYMMDD Thema ENG.jpg`
- `YYMMDD Thema FR.jpg`

Die sechs Ziffern am Anfang sind das Freigabedatum. Der Dateiname ist für die
Automatik maßgeblich. Dateien mit einem späteren Datum werden noch nicht
veröffentlicht. Bereits fällige Dateien werden bei jedem Lauf erneut geprüft und
bei Änderungen aktualisiert.

## Ausführung

Der Workflow `Bibel-Worte veröffentlichen` läuft samstags um 18:00 Uhr in der
Zeitzone `Europe/Berlin`. Er kann zusätzlich unter **Actions** manuell gestartet
werden.

## Einmalige Berechtigung

Im öffentlichen Repository wird das Action-Secret `INPUT_REPO_TOKEN` benötigt.
Es enthält einen fein eingeschränkten GitHub-Zugangsschlüssel mit ausschließlich
lesendem Zugriff auf das private Repository `bibel-worte-eingang`.

Das eingebaute `GITHUB_TOKEN` schreibt die erzeugten Dateien ausschließlich in
dieses öffentliche Repository.

## Veröffentlichungsergebnis

Die veröffentlichten Dateien liegen unter:

- `published/manifest.json`
- `published/statusbilder/<thema>/...`

Ältere Veröffentlichungen bleiben bestehen. Ein Thema mit derselben Thema-ID
wird durch die neueste fällige Fassung aktualisiert.
