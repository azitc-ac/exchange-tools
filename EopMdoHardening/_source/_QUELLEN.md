# Quelldaten für diesen Ordner

## Was hier liegt

| Datei | Zweck |
|---|---|
| `links.py` | Die recherchierten Fundorte: je Prüfpunkt Klickpfad, Portal-Deep-Link und Microsoft-Learn-Abschnitt. Dazu die Kürzel-Legende (`PREFIXES`), die skriptgenerierten IDs (`ALIASES`), die nicht von Microsoft dokumentierten EAC-Links (`UNDOCUMENTED`) und die vier Prüfpunkte ohne passgenauen Learn-Artikel (`WEAK_LEARN`). |
| `build_readme.py` | Erzeugt `README.md` aus `links.py`. `python3 build_readme.py` im Ordner. |

`links.py` ist die einzige Stelle, an der die Fundorte gepflegt werden. Daraus speisen sich die
README, die Spalten *Fundort* / *Portal öffnen* / *Microsoft Learn* in der Checkliste, Anhang E
im Guide und die Tabelle `$script:Links` im Audit-Skript.

Alle 120 Learn-Sprungmarken sind gegen die Quelldateien von `MicrosoftDocs/defender-docs` und
`MicrosoftDocs/office-docs-powershell` geprüft (Stand August 2026). Die Defender-Deep-Links
stammen aus den Learn-Artikeln selbst. Nicht belegt sind fünf EAC-Links, siehe `UNDOCUMENTED`.

## Was in diesem Ordner noch fehlt

Diese Dateien gehören dazu, sind aber beim Abräumen der Cloud-Sitzung verloren gegangen und
liegen nur noch im ausgelieferten `EOP-MDO_Hardening-Standard.zip`:

- `Invoke-EopAudit.ps1` — rein lesendes Audit, 120 Prüfpunkte, CSV + HTML-Report + JSON-Export
- `Invoke-EopHardening.ps1` — setzt die Baseline, Vorschau ist Standardmodus
- `EOP-MDO_Assessment-Checkliste.xlsx` — 120 Prüfpunkte zum Ausfüllen, Blatt „Kürzel" als Legende
- `EOP-MDO_Best-Practice-Guide.docx` / `.pdf` — Begründung je Prüfpunkt, Anhang E ist die Fundort-Tabelle

Nicht mehr vorhanden sind außerdem die Bauquellen des Guides und der Checkliste
(`items_a..d.py`, `guide_*.md`, `build_xlsx.py`, `build_guide.py`, `build_appendix_e.py`,
`verify_links.py`) sowie die Testskripte (`Mock-Exo.ps1`, `Test-Ps51Compat.ps1`,
`Test-Connection.ps1`, `Test-Report.ps1`). Wer den Guide künftig ändern will, braucht die wieder.

## Konventionen, die eingehalten werden müssen

- PowerShell-Dateien: UTF-8 **mit BOM**, kompatibel zu Windows PowerShell 5.1
- `.gitattributes` des Repos setzt `* -text` — keine Normalisierung, BOM und CRLF bleiben
- IDs sind über Checkliste, Guide und beide Skripte identisch; `links.py` kennt sie alle
