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

## Pfade/Config (konfigurierbar)

Du kannst den Pfad zur `config.json` jetzt **frei (UNC eingeschlossen)** vorgeben:

**Priorität:**
1. **Parameter** an `MainGUI5.ps1` / EXE:
   - `-ConfigPath "UNC\oder\lokal\config.json"`
   - `-ConfigDir "\\server\share\folder"`
2. **Umgebungsvariablen**:
   - `ADDUSER_CONFIG_PATH`
   - `ADDUSER_CONFIG_DIR`
3. **Fallback**: automatisch relativ zur EXE/Repo (`dist\config\config.json` bzw. `<repo>\config\config.json`)

Beispiele:

```powershell
.\MainGUI5.ps1 -ConfigPath "\\office.dir\files\ORG\OrgDATA\IT-BMU\03_Tools\AddUser-GUI\AddUser_v22\config.json"

$env:ADDUSER_CONFIG_PATH="\\srv\share\cfg\config.json"
.\MainGUI5.ps1
```
