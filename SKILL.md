---
name: solidworks-tubest-extract
description: >-
  Automates extracting hollow structural tubes and profiles from SolidWorks
  cabinets, racks, and frames (.SLDPRT weldments, .SLDASM assemblies, or .STEP
  models) into verified single-solid .STEP files, a combined tubes-only .SAT,
  and a cutting BOM for Bochu TubesT laser tube cutting. Use whenever the user
  asks to 拆图, 拆管, extract tubes from a cabinet/frame, or prepare CAD models
  for TubesT.
---

# SolidWorks Auto Tube Extraction for TubesT (激光切管拆图)

## Requirements & Paths
- **SolidWorks** (Windows desktop installation with `SolidWorks.Interop.sldworks.dll` and `SolidWorks.Interop.swconst.dll`, auto-detected by the script or passed via `-SolidWorksDir`).
- **Bochu TubesT** (for importing the exported `.step` or `.sat` tube models).
- **Default Output Root**: `%USERPROFILE%\Desktop\拆图\<ProjectName>\`
  - Always save extracted outputs into a clean subfolder under the user's Desktop `拆图` directory unless the user explicitly specifies another path.

## Mandatory Technical Rules
1. **Never use late-bound PowerShell COM** (`New-Object -ComObject SldWorks.Application` fails with `0x8002802B`). Always use C# `Add-Type` referencing the SolidWorks Interop DLLs, and preload both DLLs via `[System.Reflection.Assembly]::LoadFrom(...)` before `Add-Type`.
2. **Filter non-tube geometry before export**:
   - Multi-body weldment `.SLDPRT`: Group by `CutListFolder` for exact deduplication and quantities; exclude thin flat plates/gussets/end-caps (`minDim < 12 mm` and `FaceCount <= 12`).
   - Assemblies `.SLDASM` / imported `.STEP`: Exclude fasteners, brackets, hinges, and panels; keep structural tubes/profiles and aggregate counts by `(PartPath, ReferencedConfiguration)`.
3. **Metric naming & dual export**:
   - Name single-tube exports with metric dimensions in mm and quantity (e.g., `<Prefix>_T01_50.8x76.2_L1276mm_2pcs.step`).
   - Always generate:
     - Individual single-tube `.step` files
     - Combined pure-tube `*_AllTubesOnly.sat` (preferred by TubesT for whole-frame import)
     - `*_BOM.csv` (UTF-8 with BOM so Chinese Excel opens cleanly)
     - `*_00_Original.png` and `*_00_TubesOnly.png` visual verification snapshots
4. **Post-export validation (Required)**:
   - Verify every exported `.step` file contains **exactly 1** `MANIFOLD_SOLID_BREP` solid.
   - Cleanly close any SolidWorks instances spawned by the script so they do not consume RAM or hang on modal dialogs.
5. **TubesT handoff**:
   - TubesT requires user login (WeChat QR / account) on the user's interactive desktop. Never attempt to bypass login or launch duplicate TubesT processes if `TubesT` is already running; notify the user that verified files are ready in `%USERPROFILE%\Desktop\拆图\<ProjectName>\`.

## Workflow Steps
1. Run the helper script [extract_tubes.ps1](./scripts/extract_tubes.ps1):
   ```powershell
   powershell -ExecutionPolicy Bypass -File "<skill_dir>\scripts\extract_tubes.ps1" -InputPath "<path\to\model>" -ProjectName "<ProjectName>"
   ```
2. Verify all exported `.step` files report `OK (1 solid)` and inspect `*_00_Original.png` vs `*_00_TubesOnly.png`.
3. Provide clickable `file://` links to the output folder and BOM CSV.
