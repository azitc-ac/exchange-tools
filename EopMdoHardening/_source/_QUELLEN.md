# Quelldaten für diesen Ordner

## Was hier liegt

| Datei | Zweck |
|---|---|
| `links.py` | Die recherchierten Fundorte: je Prüfpunkt Klickpfad, Portal-Deep-Link und Microsoft-Learn-Abschnitt. Dazu die Kürzel-Legende (`PREFIXES`), die skriptgenerierten IDs (`ALIASES`), die nicht von Microsoft dokumentierten EAC-Links (`UNDOCUMENTED`) und die vier Prüfpunkte ohne passgenauen Learn-Artikel (`WEAK_LEARN`). |
| `build_readme.py` | Erzeugt `README.md` aus `links.py`. `python3 build_readme.py` im Ordner. |
| `Test-Links.ps1` | Ersatz für das verlorene `verify_links.py`. Offline: Legende, Fundort-Zeilen und URL-Menge von README und Audit-Skript stimmen überein. Mit `-Online`: jede Learn-Seite antwortet, jeder Sprunganker existiert. Portal-Links verlangen Anmeldung und werden nicht geprüft. |

`links.py` ist die einzige Stelle, an der die Fundorte gepflegt werden. Daraus speisen sich die
README, die Spalten *Fundort* / *Portal öffnen* / *Microsoft Learn* in der Checkliste, Anhang E
im Guide und die Tabelle `$script:Links` im Audit-Skript.

Alle 120 Learn-Sprungmarken sind gegen die Quelldateien von `MicrosoftDocs/defender-docs` und
`MicrosoftDocs/office-docs-powershell` geprüft (Stand August 2026). Die Defender-Deep-Links
stammen aus den Learn-Artikeln selbst. Nicht belegt sind fünf EAC-Links, siehe `UNDOCUMENTED`.

## Herkunft der ausgelieferten Dateien

Checkliste, Guide (`.docx` und `.pdf`) und die beiden Skripte stammen aus
`EOP-MDO_Hardening-Standard_5.zip` (Dateien vom 04.09.2026). Die Skripte wurden danach im
Repo angepasst (Prolog, siehe `build\`); Checkliste und Guide sind unverändert.

## Was in diesem Ordner noch fehlt

Nicht mehr vorhanden sind die Bauquellen des Guides und der Checkliste
(`items_a..d.py`, `guide_*.md`, `build_xlsx.py`, `build_guide.py`, `build_appendix_e.py`)
sowie die Testskripte (`Mock-Exo.ps1`, `Test-Ps51Compat.ps1`,
`Test-Connection.ps1`, `Test-Report.ps1`). Wer den Guide künftig ändern will, braucht die wieder.

## Konventionen, die eingehalten werden müssen

- PowerShell-Dateien: UTF-8 **mit BOM**, kompatibel zu Windows PowerShell 5.1
- `.gitattributes` des Repos setzt `* -text` — keine Normalisierung, BOM und CRLF bleiben
- IDs sind über Checkliste, Guide und beide Skripte identisch; `links.py` kennt sie alle
