# adduser_new

Refactor der **Struktur**, **kein funktionales Verhalten geändert**.  
MainGUI5.ps1 bleibt im Root. Dein bestehender Code (Signaturen, Pfade, ProcessUserCreation, etc.) bleibt wie er ist.

## Struktur (neu)

```
adduser_new/
├─ MainGUI5.ps1                     # unverändert im Root
├─ AddUser.psd1 / AddUser.psm1      # Root-Modul (lädt alle Module), optional zu benutzen
├─ Modules/
│  ├─ Public/
│  └─ Internal/
├─ config/
├─ Xaml/
├─ scripts/
│  ├─ build.ps1                     # EXE-Build via PS2EXE (optional)
│  └─ Invoke-Tests.ps1              # Pester starten
├─ tests/                           # Pester-Tests (minimal, nicht-invasiv)
└─ .github/workflows/ci.yml         # Pester im CI
```

## GUI starten (wie gehabt)

```
.\MainGUI5.ps1
```

## Optional: Root-Modul verwenden

Wenn du lokal mit den Modulen arbeiten willst (ohne UNC-Importe), kannst du

```powershell
Import-Module .\AddUser.psd1 -Force
```

nutzen – der bestehende Code ändert sich dadurch nicht.

## Tests

```powershell
pwsh -f scripts\Invoke-Tests.ps1
```

## EXE bauen (optional, ohne Funktionalitätsänderung)

```powershell
pwsh -f scripts\build.ps1
```

Ergebnis: `dist\AddUserGUI.exe` (+ Payload-Ordner).

## CI

Der Workflow `.github/workflows/ci.yml` läuft Pester bei jedem Push/PR.

## Pfade/Config

Die existierende Config/Pfad-Logik bleibt wie sie ist (UNC-Pfade etc.).  
Wenn du später dynamischere Pfadauflösung willst, können wir das separat **ohne** Änderung der jetzt gelieferten Struktur einbauen (z.B. per Env‑Vars, CLI‑Param oder auto-resolve relativ zur EXE).
