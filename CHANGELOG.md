# Changelog

## [4.0.0] – 2026-11-27

### Hinzugefügt
- **Delta-Update** des Caches über das Audit-Log (`Search-UnifiedAuditLog`)
  - liest nur Postfächer nach, deren Vollzugriff, Kalenderrechte oder Buchungsstellvertreter geändert wurden
  - erkennt neue, gelöschte und umgewandelte Postfächer
  - im Dialog „Cache aktualisieren“, im neuen Menü **Cache** und in der Konsole (`-DeltaUpdate`)
  - Ergebnisprotokoll `delta.log`; benötigt die Rolle *View-Only Audit Logs*
- **Update-Hinweis**: Prüfung auf neue GitHub-Releases beim Start (Hintergrund), Link in der Statusleiste
- Menü **Hilfe**: Nach Updates suchen, Projekt auf GitHub öffnen, Über (Version, Rollen, Datenordner)
- Parameter `-NoUpdateCheck`
- **GitHub Actions**: CI-Prüfung (Syntax, UTF-8-BOM, PSScriptAnalyzer) und Release-Workflow (Tag → EXE + Release + SHA256)
- README mit Screenshots-Bereich und englischer Zusammenfassung

### Geändert
- Quellcode aufgeteilt in `src\*.ps1`; `Get-UserMailboxAccess.ps1` wird mit `build\Merge-Source.ps1` zusammengesetzt
- `Build-MailboxAccessExe.ps1` setzt die Quellen vor dem Build automatisch zusammen und übernimmt die Version aus dem Skript
- Cache-Alter zählt ab Komplettaufbau **oder** letztem Delta-Update
- Fenstertitel zeigt die Version

## [3.9.0] – 2026-11-27

### Hinzugefügt
- Menü **Extras**:
  - **Weiterleitungen & Posteingangsregeln** des Benutzers – externe Weiterleitungen rot markiert
  - **Gruppenmitgliedschaften** anzeigen, hinzufügen, entfernen (Verteiler, E-Mail-Sicherheitsgruppen, Microsoft 365)
  - **Ordnerrechte** eines Postfachs (alle sichtbaren Ordner) – auch im Rechtsklick-Menü und im Fenster „Postfach-Zugriff“
  - **Tenant-Gesamtreport**: alle Postfachrechte + Weiterleitungen als HTML (filter- und sortierbar) und CSV
  - Änderungsprotokoll, Datenordner öffnen
- Gruppenänderungen werden im `aenderungen.log` protokolliert

### Geändert
- Nach **Erteilen / Entfernen** wird die Änderung sofort in die Liste übernommen
  (Exchange liefert „Senden als“ / „Senden im Auftrag von“ teils verzögert)
  - automatische Nachprüfung der Liste nach 45 Sekunden
  - offenes Fenster „Postfach-Zugriff“ wird ebenfalls aktualisiert

## [3.8.1] – 2026-11-27

### Behoben
- Fehler beim Start bzw. in der Oberfläche führten zum stillen Beenden ohne Meldung
  - jetzt Meldungsfenster + `start_fehler.log` / `ui_fehler.log` in `%LOCALAPPDATA%\EXOMailboxAccess`
- Autovervollständigung wird nur noch neu zugewiesen, wenn sich die Postfachliste ändert

## [3.8.0] – 2026-11-27

### Hinzugefügt
- **Rechte kopieren** (Onboarding „wie Kollege“): direkte Rechte des abgefragten Benutzers auf einen anderen übertragen
  - Auswahl pro Recht, AutoMapping optional, Kalenderrolle wird übernommen
- **Postfach-Zugriff**: eigenes Fenster „Wer hat Zugriff auf Postfach X?“
  - Vollzugriff, Kalender, Buchungsstellvertreter aus dem Cache; Senden als / im Auftrag live
  - Kopieren, CSV-Export, Doppelklick auf Benutzer öffnet dessen Rechte
  - Doppelklick in der Hauptliste öffnet das Fenster für das Postfach
- **Offboarding**: alle direkten Rechte eines Benutzers entfernen, vorher automatische Nachweis-CSV unter `%LOCALAPPDATA%\EXOMailboxAccess\Offboarding`
- **Autovervollständigung** für Benutzer- und Postfachfelder (Name oder Adresse)
- **Zuletzt abgefragte Benutzer** (Button ▾ neben dem Benutzerfeld, max. 15)
- **Protokoll-Fenster** für `aenderungen.log` mit Filter und Farbmarkierung
- **Rechteprüfung beim Start**: ohne Exchange-Rolle zum Ändern läuft das Tool im Nur-Lesen-Modus

### Geändert
- Kalenderberechtigungen können jetzt auf allen Postfachtypen erteilt werden (Voraussetzung für das Kopieren)

## [3.7.0] – 2026-11-27

### Hinzugefügt
- Button **Rechte erteilen...** (grün), auch im Rechtsklick-Menü
- Dialog mit Postfachsuche, Mehrfachauswahl per Haken und Filter „nur Raum-/Geräte-Postfächer“
  - in der Liste markierte Postfächer sind vorausgewählt
- Erteilbare Rechte:
  - Lesen und Verwalten (Vollzugriff), optional ohne AutoMapping
  - Senden als
  - Senden im Auftrag von
  - nur Raum/Gerät: Kalenderberechtigung (Rolle wählbar), Buchungsstellvertreter
- Erteilte Rechte werden sofort in den Cache übernommen und in `aenderungen.log` protokolliert (`ERTEILT`)

## [3.6.0] – 2026-11-27

### Hinzugefügt
- **Raum- und Gerätepostfächer**:
  - Kalenderberechtigungen werden für Raum-/Gerätepostfächer immer eingelesen, auch ohne „Kalenderrechte im Cache“
  - neue Berechtigung **Buchungsstellvertreter** (ResourceDelegates aus `Get-CalendarProcessing`)
- Haken **Raumpostfächer anzeigen** (Standard: an)
- **Rechte entfernen** kann jetzt auch:
  - **Kalenderberechtigung** (`Remove-MailboxFolderPermission`)
  - **Buchungsstellvertreter** (`Set-CalendarProcessing -ResourceDelegates`)
- Filter „Buchungsstellvertreter“ in der Typ-Auswahl
- „Live prüfen“ liest bei Räumen zusätzlich Kalender und Stellvertreter nach

### Geändert
- Ein Cache ohne Raumdaten (vor 3.6) gilt als veraltet, beim Start erscheint die Abfrage zum Neuaufbau

## [3.5.0] – 2026-11-27

### Hinzugefügt
- Button **Live prüfen**: liest den Vollzugriff (und ggf. Kalender) live nach für
  - alle aktuell angezeigten Postfächer
  - alle Shared-, Raum- und Geräte-Postfächer
  - aktualisiert den Cache gezielt, ohne kompletten Neuaufbau
- Dialog **Cache aufbauen** mit Haken **Kalenderrechte mit einlesen**
- Beim Start: Abfrage, ob ein veralteter Cache neu aufgebaut werden soll
- Tooltips für Cache-Aktionen

### Geändert
- Standard-Cache-Alter: **1 h** (vorher 24 h); Konsolenmodus baut weiterhin erst ab 24 h automatisch neu
- Checkbox „inkl. Kalenderrechte“ heißt jetzt „Kalenderrechte im Cache“

## [3.4.0] – 2026-11-27

### Hinzugefügt
- Button **Rechte entfernen...** (Zeile „Ansicht“) und Rechtsklick-Menü in der Liste
  - entfernt **direkte** FullAccess-, SendAs- und SendOnBehalf-Rechte des abgefragten Benutzers
  - Dialog mit Checkboxen je Berechtigung (Bezeichnungen wie im Exchange Admin Center:
    *Lesen und Verwalten (Vollzugriff)*, *Senden als*, *Senden im Auftrag von*) inkl. Anzahl, Vorschau und Sicherheitsabfrage
  - Rechte über Gruppen und vererbte Rechte werden angezeigt, aber nicht verändert
  - Ergebnisbericht (OK/FEHLER je Postfach), Liste wird danach automatisch neu geladen
  - FullAccess-Einträge werden direkt aus dem Cache entfernt
  - Protokoll: `%LOCALAPPDATA%\EXOMailboxAccess\aenderungen.log`
- Rechtsklick-Menü: **Adresse(n) kopieren**

### Geändert
- **Rechte über Gruppen anzeigen** ist standardmäßig aus (nur direkte Rechte)

## [3.3.0] – 2026-11-27

### Geändert
- Spaltenreihenfolge: **Postfach** zuerst, danach Berechtigung(en) – auch im CSV-Export

### Hinzugefügt
- App-Icon in der Titelleiste und in der Taskleiste (statt PowerShell-Logo)
  - eigene AppUserModelID `Example.MailboxAccess`
  - Icon wird neben dem Skript gesucht (`mailbox.ico`)
- Build bettet das Icon ein und legt es beim Start nach `%LOCALAPPDATA%\EXOMailboxAccess\`
- Build nimmt `mailbox.ico` automatisch, wenn kein `-IconFile` angegeben ist

## [3.2.0] – 2026-11-27

### Hinzugefügt
- Checkbox **Jedes Postfach nur einmal (Rechte zusammenfassen)** – Standard: an
  - eine Zeile pro Postfach, Berechtigungen kombiniert (z. B. `FullAccess, SendAs`)
  - Quelle kombiniert, `direkt` zuerst (z. B. `direkt; Gruppe: GRP-Admin`)
  - Spalte „Rechte“ wird in dieser Ansicht ausgeblendet
- Neue Zeile „Ansicht“ im Kopfbereich (Zusammenfassen + Gruppen-Checkbox)

### Geändert
- Dropdown-Filter arbeitet mit LIKE (funktioniert in beiden Ansichten)
- Suchfeld durchsucht jetzt auch die Spalte Berechtigung(en)
- CSV-Export und „Auswahl kopieren“ übernehmen die aktuelle Ansicht (nur sichtbare Spalten)

## [3.1.0] – 2026-11-27

### Hinzugefügt
- Checkbox **Rechte über Gruppen anzeigen**: blendet Berechtigungen über Gruppen (z. B. GRP-Admin) ein/aus
- Statusleiste zeigt Anzahl direkter Rechte und Rechte über Gruppen

## [3.0.1] – 2026-11-27

### Behoben
- SendAs-Abfrage: `-Trustee` nutzt jetzt die ExternalDirectoryObjectId (GUID) statt der SMTP-Adresse
  (Fehler: *Trustee parameter can be used only with Identity or ExternalDirectoryObjectId or UserPrincipalName*)
- Fallback auf `Get-RecipientPermission`, falls `Get-EXORecipientPermission` fehlschlägt
- Build: Skript wird immer als UTF-8 mit BOM eingebettet (behebt kaputte Umlaute wie „PostfÃ¤cher“)

## [3.0.0] – 2026-11-27

### Hinzugefügt
- GUI läuft in eigenem STA-Runspace – Fenster bleibt während Cache-Aufbau bedienbar (kein „Keine Rückmeldung“ mehr)
- Button **Abbrechen** für laufenden Cache-Aufbau (bisheriger Cache bleibt erhalten)
- Laufbalken-Animation während des Ladens der Postfachliste
- Abfrage beim Schließen, wenn noch ein Vorgang läuft
- `Build-MailboxAccessExe.ps1`: EXE als Starter mit eingebettetem Skript

### Geändert
- EXE führt das Skript im echten `powershell.exe` aus statt im PS2EXE-Host

## [2.1.0] – 2026-11-27

### Hinzugefügt
- Automatische Wiederholung bei Verbindungsabbrüchen (Postfachliste 4×, je Postfach 3×)
- Fehlgeschlagene Postfächer werden protokolliert (`cache_fehler.log`) und in der GUI angezeigt
- TLS 1.2 wird erzwungen

### Behoben
- WAM-Anmeldung ohne Konsolenfenster (`A window handle must be configured`)
- `#Requires` durch Modulprüfung ersetzt
- Doppelte BOM am Dateianfang

## [2.0.0] – 2026-11-27

### Hinzugefügt
- Berechtigungs-Cache für FullAccess und Kalender (`%LOCALAPPDATA%\EXOMailboxAccess\cache.xml`)
- Auflösung von Berechtigungen über mail-aktivierte Gruppen (Spalte „Quelle“)
- Cache-Alter und -Status in der GUI

## [1.0.0] – 2026-11-27

### Hinzugefügt
- Abfrage von FullAccess, SendAs, SendOnBehalf und Kalenderrechten
- WinForms-GUI mit Filter, Farbmarkierung, CSV-Export und Zwischenablage
- Konsolenmodus (`-NoGui`) mit gruppierter Ausgabe
