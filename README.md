# Lebensmitteleinkauf

Responsive Web-App für die antientzündliche Lebensmittel- und Einkaufsliste.

## Start

Die App benötigt keinen Build-Schritt. `index.html` kann direkt geöffnet oder über
einen beliebigen lokalen Webserver bereitgestellt werden.

## Funktionen

- 269 Lebensmittel aus der Excelvorlage mit Suche und Filtern
- lokal gespeicherte Einkaufsliste mit optionaler OneDrive-Synchronisierung
- TXT-Export, Kopieren und Zurücksetzen der Markierungen
- Tagesbaukasten und dynamische Auswertung
- eigenständige Layouts für große und kleine Bildschirme

## OneDrive-Synchronisierung

Die App bleibt eine statische GitHub-Pages-App und nutzt MSAL.js im Browser.
Es wird kein Client Secret verwendet. Angemeldete Nutzer speichern ihre Liste in
ihrem eigenen OneDrive-App-Ordner mit `Files.ReadWrite.AppFolder`.

Azure App Registration:

- Client-ID: `3c4004e7-9323-440c-8977-96699d8d8e6f`
- produktive Redirect-URI: `https://torwoe.github.io/lebensmitteleinkauf/`
- lokale Redirect-URI fuer Tests mit einem einfachen lokalen Webserver:
  `http://localhost:5500/`

Wenn statt `localhost` die IP-Adresse genutzt wird, muss diese Redirect-URI
separat in Azure eingetragen werden: `http://127.0.0.1:5500/`.
