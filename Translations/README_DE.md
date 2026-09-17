# PS-NCDU

**Festplattenplatz-Analyse für Windows, in PowerShell, mit lokaler Weboberfläche und in Echtzeit navigierbarem Verzeichnisbaum.**

[![Version](https://img.shields.io/badge/version-6.28-2c6cb0)](https://github.com/Vietnamix/PS-NCDU)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell&logoColor=white)](https://github.com/Vietnamix/PS-NCDU)
[![Plattform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)](https://github.com/Vietnamix/PS-NCDU)
[![Lizenz](https://img.shields.io/badge/license-MIT-3fa45b)](License.md)

PS-NCDU ist ein **eigenständiges Ein-Datei**-PowerShell-Skript, das in Sekunden die Frage beantwortet: „Was füllt diese Festplatte?“. Ohne Parameter gestartet, öffnet es einen kleinen lokalen Webserver, startet den Browser und lässt Sie einen Ordner oder ein Laufwerk zur Analyse wählen. Der Baum entsteht **live** während des Scans, die Größen füllen sich Ordner für Ordner, und Sie können frei navigieren, noch bevor der Scan beendet ist. Inspiriert vom Unix-Werkzeug [`ncdu`](https://dev.yorhel.nl/ncdu), für das Windows-Ökosystem gebaut, ohne jede externe Abhängigkeit.

![PS-NCDU-Oberfläche](PS-NCDU_interface_v6.27b.png)

*PowerShell 5.1+ · Windows · Keine Installation · Eine Datei · 42 Sprachen*

Weitere Sprachen: [Français](README.md) · [English](README_EN.md) · [中文](README_ZH.md) · [हिन्दी](README_HI.md) · [Español](README_ES.md) · [العربية](README_AR.md) · [বাংলা](README_BN.md) · [Português](README_PT.md) · [Русский](README_RU.md) · [اردو](README_UR.md) · [Bahasa Indonesia](README_ID.md) · [日本語](README_JA.md) · [Türkçe](README_TR.md) · [Tiếng Việt](README_VI.md) · [한국어](README_KO.md) · [Italiano](README_IT.md)

---

## Inhalt

- [Überblick](#überblick)
- [Funktionen](#funktionen)
- [Voraussetzungen](#voraussetzungen)
- [Installation](#installation)
- [Verwendung](#verwendung)
- [Die Oberfläche](#die-oberfläche)
- [So arbeitet der Scan](#so-arbeitet-der-scan)
- [Arbeitsdateien](#arbeitsdateien)
- [Fehlerbehebung](#fehlerbehebung)
- [Roadmap](#roadmap)
- [Mitwirken](#mitwirken)
- [Lizenz](#lizenz)

---

## Überblick

Anders als die 3.x-Versionen, die einen statischen HTML-Bericht zum späteren Öffnen erzeugten, ist PS-NCDU heute eine **lokale Webanwendung**. Das Skript startet einen HTTP-Server auf `127.0.0.1` (Port 8787, mit automatischem Ausweichen auf einen freien Port), geschützt durch ein Sitzungstoken, und öffnet die Oberfläche im Standardbrowser. Alles wird in dieser Oberfläche eingestellt: der zu analysierende Ordner, die Tiefe, der Anzeigefilter, die Ausschlüsse, die Sprache.

Der Server bleibt auf dem lokalen Rechner, ist im Netzwerk nicht erreichbar und wird mit `Strg+C` in der Konsole oder über die Schaltfläche „Server beenden“ in der Oberfläche gestoppt.

![PS-NCDU-Analysefenster](PS-NCDU_scan_form_v6.27b.png)

---

## Funktionen

### Scan und Navigation
- **Echtzeit-Baum**: Die Struktur erscheint während der Aufzählung, die Größen kommen Ordner für Ordner, sobald sie berechnet sind.
- **Freie Navigation während des Scans**: Klicken Sie auf einen Ordner, um ihn zu öffnen, und auf den Brotkrumenpfad, um nach oben zu gehen, ohne auf das Ende zu warten.
- **Einstellbare oder unbegrenzte Tiefe**: einige Ebenen vorladen für eine flüssige Anzeige, oder den ganzen Baum. Die Größen sind unabhängig von der Tiefe immer exakt; nicht vorgeladene Ebenen laden per Klick.
- **Verschränkter unbegrenzter Scan**: Bei unbegrenzter Tiefe wird jeder Teilbaum unmittelbar vor der Messung aufgezählt, sodass Größen schon nach Sekunden erscheinen, statt auf den Durchlauf der ganzen Festplatte zu warten.
- **Abbruch**: Eine Schaltfläche „Abbrechen“ stoppt den laufenden Scan und gibt die Kontrolle zurück; ein neuer Scan bricht den alten automatisch ab.
- **Anklickbare graue Ordner**: ausgeschlossene, Junctions, geschützte (ACL) oder nicht vorgeladene Ordner bleiben sichtbar und werden auf Wunsch gescannt, mit Warteschlange, falls bereits ein Scan läuft.
- **Verarbeitungsreihenfolge passend zur Anzeige**: Der Fortschritt füllt sich von oben nach unten, ohne Sprünge.

### Ergebnisse lesen
- **Sortierung Name / Größe**: standardmäßig nach Größe (die größten oben), während des Scans geglättet, damit Zeilen nicht springen; ein Klick wechselt zur Sortierung nach Name.
- **Anteilsbalken** und Prozentwerte relativ zum aktuellen Ordner.
- **Rekursive Zähler** für Unterordner und Dateien je Ordner.
- **Dateien auf Abruf**, wenn Sie einen Ordner öffnen, nach Größe sortiert, begrenzt auf die 1000 größten.
- **Symbole je Dateityp**: rund 120 gängige Endungen (Bilder, Video, Audio, PDF, Office, Archive, Code, ausführbare Dateien, Schriften, Datenträgerabbilder, Datenbanken, E-Books, Zertifikate, Verknüpfungen), um Typen auf einen Blick zu erkennen.
- **Anzeigefilter**: Elemente unter 1 MB, 100 MB oder 1 GB ausblenden, für die Lesbarkeit, ohne den Scan zu verändern.
- **Farbige Statuspunkte**: gescannt, geplant, in der Warteschlange, in Arbeit, grau; eine integrierte Legende und Hilfe erklären jeden Zustand.
- **Helles / dunkles Design**.

### Analysefenster
- **Integrierter Ordnerbrowser**: Laufwerke, Navigation per Klick, übergeordneter Ordner, „Diesen Ordner wählen“. Keine Abhängigkeit vom nativen Windows-Dialog, daher auch per Fernzugriff zuverlässig.
- **Schnellzugriff** auf Benutzerprofile, **Zuletzt verwendet** mit Einzelentfernung und Leeren, **Laufwerke** mit Belegungsbalken.
- **Pfadprüfung** live und beim Start.
- **Ausschlüsse**: Liste der immer übersprungenen Systemordner, plus ein Feld, um für einen Scan weitere auszuschließen.
- **Gemerkte Einstellungen** (Pfad, Tiefe, Filter, Sortierung, Sprache) über Sitzungen hinweg.
- **Eingabe** zum Starten, **Esc** oder Schließen-Schaltfläche, um das Fenster auszublenden, wenn bereits ein Scan angezeigt wird.

### Sprachen
- **42 Sprachen**, die über 80 % der Weltbevölkerung abdecken: Englisch, Chinesisch, Hindi, Spanisch, Französisch, Arabisch, Bengalisch, Portugiesisch, Russisch, Urdu, Indonesisch, Deutsch, Japanisch, Koreanisch, Italienisch, Türkisch, Vietnamesisch, Polnisch, Niederländisch, Ukrainisch, Rumänisch, Tschechisch, Griechisch, Schwedisch, Ungarisch, Persisch, Thai, Malaiisch, Filipino, Suaheli, Tamil, Telugu, Marathi, Gujarati, Kannada, Malayalam, Panjabi, Hebräisch, Hausa, Birmanisch, Amharisch, Khmer.
- Automatische Erkennung der Systemsprache, Auswahl im Analysefenster, gemerkte Wahl.
- Schreibrichtung von rechts nach links für Arabisch, Urdu, Persisch und Hebräisch.
- Die Übersetzungen der 30 jüngsten Sprachen sind nach bestem Bemühen erstellt; eine Durchsicht durch Muttersprachler ist willkommen, besonders für Amharisch, Khmer, Birmanisch, Hausa und die indischen Sprachen.

---

## Voraussetzungen

| Element    | Details                                                                                                                             |
| ---------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| System     | Windows 10 / 11 oder Windows Server                                                                                                 |
| PowerShell | 5.1 (Windows PowerShell) oder 7+ (PowerShell Core)                                                                                  |
| Modus      | **FullLanguage** erforderlich (der Webserver nutzt `HttpListener`). Siehe [Fehlerbehebung](#fehlerbehebung) bei eingeschränktem Modus. |
| Rechte     | Lesezugriff auf die gescannten Ordner; manche Systempfade erfordern eine **Administrator**-Konsole                                  |
| Browser    | Jeder aktuelle Browser                                                                                                              |

Es wird kein externes Modul benötigt.

---

## Installation

Klonen Sie das Repository oder laden Sie einfach die Datei `ps-ncdu.ps1` herunter:

```powershell
git clone https://github.com/Vietnamix/PS-NCDU.git
cd PS-NCDU
```

Das Skript ist in **UTF-8 mit BOM** kodiert. Speichern Sie es nicht in einer anderen Kodierung: PowerShell 5.1 würde die Datei sonst als ANSI lesen und die Akzente sowie die nicht-lateinischen Sprachen der Oberfläche zerstören.

> **Ausführungsrichtlinie**: Wenn Windows die Ausführung von Skripten blockiert, erlauben Sie sie für die aktuelle Sitzung:
>
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> ```
>
> Dieser Befehl ändert nichts dauerhaft: Er gilt nur für das geöffnete PowerShell-Fenster.

---

## Verwendung

Es gibt **keine Kommandozeilenparameter**. Starten Sie einfach das Skript:

```powershell
.\ps-ncdu.ps1
```

Das Skript:

1. startet den lokalen Server und zeigt seine Adresse in der Konsole an (zum Beispiel `http://127.0.0.1:8787/?token=...`);
2. öffnet diese Adresse im Standardbrowser;
3. zeigt das Analysefenster, in dem Sie Ordner, Tiefe und Anzeige wählen und dann auf „Analysieren“ klicken.

Öffnet sich der Browser nicht von selbst, kopieren Sie die in der Konsole angezeigte Adresse. Zum Beenden: `Strg+C` in der Konsole oder die Schaltfläche „Server beenden“ in der Oberfläche.

Um Systempfade zu scannen (`C:\Windows`, das Stammverzeichnis eines Laufwerks), starten Sie die Konsole **als Administrator**: Geschützte Ordner erscheinen sonst grau.

---

## Die Oberfläche

- **Kopfzeile**: Brotkrumenpfad, Gesamtgröße des aktuellen Ordners mit Zählern, Fortschritt in Prozent, Sortierschaltfläche Name / Größe, Design, „Abbrechen“ während eines Scans, „Neuer Scan“.
- **Baum**: eine Zeile je Ordner oder Datei, mit Statuspunkt, Typsymbol, Name, Anteilsbalken, Prozentwert, Zählern und Größe. Graue Ordner werden per Klick gescannt; eine Schaltfläche „Graue Ordner scannen (N)“ verarbeitet alle im aktuellen Ordner.
- **Fußzeile**: aktuelle Phase, der gerade tatsächlich gelesene Pfad, Scan-Tiefe, Stoppuhr und das Panel Legende / Hilfe.
- **Analysefenster** („Neuer Scan“): zweispaltig. Links das Ziel (Pfad, integrierter Browser, Schnellzugriff, Zuletzt verwendet, Laufwerke). Rechts die Optionen (Tiefe mit Schieberegler und unbegrenztem Modus, Anzeigefilter, Ausschlüsse). Die Sprachauswahl sitzt in der Kopfzeile dieses Fensters.

---

## So arbeitet der Scan

Die Engine arbeitet in Phasen. Eine **Aufzählung** entdeckt die Struktur und streamt sie laufend in den Baum; eine **Größenberechnung** durchläuft anschließend jeden Teilbaum der ersten Ebene in Anzeigereihenfolge und reicht Teilgrößen einmal pro Sekunde an alle Vorfahren nach oben. Die Ereignisse fließen per SSE vom Server zur Seite.

Einige Entwurfsentscheidungen, die man kennen sollte:

- **Nur ein Scan zugleich**, bewusst: Zwei parallele Festplattenscans würden sich gegenseitig ausbremsen. Zusätzliche Anfragen (graue Ordner) werden eingereiht und am Ende bearbeitet.
- **Einthread-Server**: Der Abbruch ist kooperativ. Das Schließen der Verbindung (Schaltfläche „Abbrechen“ oder neuer Scan) lässt den nächsten Schreibvorgang des Servers fehlschlagen, was ein in den Schleifen geprüftes Flag setzt; der tatsächliche Stopp dauert bis zu einer Sekunde.
- **Junctions und Reparse-Punkte** werden übersprungen, um Schleifen und Doppelzählungen zu vermeiden.
- **Netzlaufwerke**: Ihr Speicherplatz wird beim Start nicht abgefragt, was ein Hängen verhindert, wenn ein zugeordnetes Laufwerk nicht erreichbar ist (zum Beispiel bei getrenntem VPN).
- **Streaming-Aufzählung über .NET** (`EnumerateFiles` / `EnumerateDirectories`) statt `Get-ChildItem`, in PowerShell 5.1 deutlich schneller. Die Leistungsgrenze bleibt die eines interpretierten Skripts: Native Werkzeuge, die die NTFS-MFT direkt lesen, sind weit schneller, und das wird bewusst hingenommen.

---

## Arbeitsdateien

| Ort                                  | Zweck                                          |
| ------------------------------------ | ---------------------------------------------- |
| `%TEMP%\psncdu\psncdu_debug.log`     | Ausführliches Protokoll von Server und Scans   |
| `%TEMP%\psncdu\history.txt`          | Verlauf der gescannten Pfade („Zuletzt verwendet“) |

Immer vom Scan ausgeschlossene Systemordner: `C:\Windows\WinSxS`, `C:\Windows\Installer`, `C:\$Recycle.Bin`, `C:\System Volume Information`, `C:\Recovery`, `C:\ProgramData\Microsoft\Windows Defender`, `C:\Windows\SoftwareDistribution`. Weitere können Sie für die Dauer eines Scans im Abschnitt Ausschlüsse des Analysefensters hinzufügen.

---

## Fehlerbehebung

| Symptom                                                        | Wahrscheinliche Ursache / Lösung                                                                                                                 |
| -------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| Das Skript startet nicht                                       | Ausführungsrichtlinie, siehe den Hinweis unter [Installation](#installation).                                                                    |
| „Der Webserver erfordert den Modus FullLanguage“               | Die Sitzung läuft in ConstrainedLanguage (AppLocker-/WDAC-Richtlinie). Starten Sie aus einer uneingeschränkten Konsole oder nutzen Sie Version 3.2 (statischer HTML-Bericht), die im eingeschränkten Modus funktioniert. |
| Akzente oder nicht-lateinische Sprachen in der Oberfläche defekt | Die `.ps1` wurde ohne UTF-8-BOM neu gespeichert. Stellen Sie die ursprüngliche Kodierung wieder her.                                            |
| Viele graue Ordner „geschützt“                                  | Unzureichende Rechte. Starten Sie PowerShell **als Administrator** neu.                                                                          |
| Der Browser öffnet sich nicht                                   | Öffnen Sie die in der Konsole angezeigte Adresse manuell (mit ihrem Token).                                                                      |
| Leere Seite oder eingefrorener Server beim Start               | Prüfen Sie das Protokoll `%TEMP%\psncdu\psncdu_debug.log`. Ein unerreichbares Netzlaufwerk konnte ältere Versionen blockieren; seit 5.14 behoben. |
| Nichts passiert bei einem unbegrenzten Scan eines Laufwerks    | Seit 6.17 behoben (verschränkte Aufzählung). Bleibt es dabei, prüfen Sie die in der Kopfzeile angezeigte Version.                                 |

---

## Roadmap

- [ ] Durchsicht der Übersetzungen der 30 jüngsten Sprachen durch Muttersprachler
- [ ] Fortschrittsbalken auf Basis des tatsächlich belegten Speicherplatzes statt auf Phasenabschnitten
- [ ] Durchsatzanzeigen während des Scans (Dateien pro Sekunde, MB pro Sekunde)
- [ ] CSV-/JSON-Export der Ergebnisse
- [ ] Vergleich zweier Scans über die Zeit

*Vorschläge sind über Issues willkommen.*

---

## Mitwirken

Beiträge sind willkommen:

1. *Forken* Sie das Repository.
2. Legen Sie einen Branch an (`git checkout -b feature/meine-funktion`).
3. Behalten Sie die Kodierung **UTF-8 mit BOM** und die PowerShell-*Here-Strings* unverändert bei.
4. Für Übersetzungen: Jede Sprache ist ein Objekt im Wörterbuch `I18N` im Skript; vergleichen Sie dessen Schlüssel mit denen von `en`, um Fehlendes zu finden.
5. Öffnen Sie einen *Pull Request* mit klarer Beschreibung der Änderung.

Für Fehler und Ideen öffnen Sie ein **Issue** mit Windows-Version, PowerShell-Version, der in der Kopfzeile angezeigten PS-NCDU-Version und nach Möglichkeit einem Auszug aus dem Protokoll `%TEMP%\psncdu\psncdu_debug.log`.

---

## Lizenz

Veröffentlicht unter der **MIT**-Lizenz. Siehe die Datei [`License.md`](License.md).

---

## Autor

**[Eric Guiffault](https://eric.guiffault.com)**

Wenn dieses Projekt Ihnen nützt, freue ich mich über einen Stern auf GitHub.
