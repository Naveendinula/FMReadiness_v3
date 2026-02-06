# FM Readiness v3

FM Readiness v3 is a Revit add-in for FM/COBie data auditing, guided fixing, bulk parameter editing, and FM export workflows.

## What it does
- Runs readiness audits for supported MEP categories using a selected preset.
- Scores readiness at element, group, category, and overall levels.
- Edits Component and Type fields from the same dockable pane.
- Supports bulk updates by current selection, type scope, active view, or category.
- Exports FM metadata to IFC-aligned outputs (`.ifc` + `.fm_params.json` sidecar).

## Recent workflow updates (now in UI)

### Audit Results
- Split summary badges into **Component Groups** and **Type Groups**.
- Added group badge tooltips (scope, included fields, checks passed, scoring mode).
- Added audit controls:
  - **Audit scope**: `All editable fields` or `Required + unique only`
  - **Audit view**: `Combined`, `Component only`, `Type only`
  - Category filter and "Show only incomplete"
- Added row actions:
  - **Fix**: switches to Parameter Editor, sets that element as working selection, and highlights audit-missing fields.
  - **Locate**: selects/zooms in Revit.
  - **View/Open**: open a 2D view context for the asset.

### Parameter Editor
- Added selection auto-sync controls:
  - **Auto-sync** (default ON)
  - **Lock selection** (default ON)
  - **Scope picker**: Current selection, Active view, All instances of selected type, By category
- Added **Selection list tray** for multi-select edits:
  - remove one element
  - remove checked elements
  - keep checked elements only
  - clear all
- Added safe-apply behavior:
  - confirmation prompts before selected/type/category writes
  - preview text showing impact counts
  - writes **only user-modified fields**
- Added field-level validation before apply:
  - ISO date normalization/validation (`YYYY-MM-DD`)
  - numeric validation for number fields
  - inline invalid styling + toast listing invalid field names
- Added **Fix mode banner** to explain highlighted fields and keep full-edit freedom.

### COBie / Legacy modes
- `Ensure COBie Parameters` now supports cleanup toggle:
  - **Remove FM_ parameters** (default ON)
- Editor still supports both:
  - **COBie Editor** (preset-driven fields)
  - **Legacy FM Editor** (FM_* fields)

## Supported categories
Audit/editor target categories:
- `OST_MechanicalEquipment`
- `OST_DuctTerminal`
- `OST_DuctAccessory`
- `OST_PipeAccessory`

Additional COBie context categories (preset-dependent):
- `OST_Rooms`
- `OST_MEPSpaces`
- `OST_Levels`

## Presets and audit profiles
Presets are in `FMReadiness_v3/Presets/`:
- `cobie-core.json`
- `fm-legacy.json`
- `custom.json`

Preset mappings define:
- Component/Type fields
- Required/unique/date rules
- alias parameters and write behavior
- group labels used in audit scoring

Audit profile state is tied to active preset + selected score mode.

## Recommended editor flow
1. Select elements in Revit (or use Scope + Apply Scope).
2. Open **Parameter Editor** and click **Refresh** if needed.
3. Edit only the fields you want changed (modified fields are tracked).
4. Resolve any invalid field warnings.
5. Apply to Selected / Update Type / Apply to Category.
6. Re-run or refresh audit to confirm score improvement.

## Audit "Fix" behavior
When you click **Fix** in Audit Results:
- the plugin targets that audit row element
- the editor highlights mapped missing fields for that row
- you can still edit any non-highlighted field

This is guidance, not a restriction.

## IFC + sidecar export behavior

### Sidecar (`*.fm_params.json`)
Current export includes:
- `FMReadiness` instance properties (FM/COBie operational fields)
- `FMReadinessType` subset (`Manufacturer`, `Model`, `TypeMark`)
- metadata block (`RevitElementId`, `RevitUniqueId`, category/family/type)

### IFC export
The add-in enables user-defined property sets from:
- `FMReadiness_v3/IFC/FMReadiness_Psets.txt`

By default, this does **not** guarantee every COBie Type editor field is in IFC.  
If you need broader IFC property coverage, include Revit property sets in IFC export setup (when prompted by the add-in / export configuration).

## Build and run

### Prerequisites
- Autodesk Revit (2022-2026 targets supported by solution configs)
- Visual Studio or Rider with .NET desktop workload
- Microsoft Edge WebView2 Runtime

### Build
1. Open `FMReadiness_v3.slnx`
2. Choose target config (example: `Debug.R26` or `Release.R24`)
3. Build solution

### Run in Revit
Use the **Digital Twin** tab / **FM Tools** panel:
- `Run FM Audit`
- `FM Pane`
- `Export FM Sidecar`
- IFC export command (if included in your ribbon setup)

## Key files
- `FMReadiness_v3/Services/AuditService.cs`
- `FMReadiness_v3/Services/CobieMappingService.cs`
- `FMReadiness_v3/UI/Panes/WebViewPaneController.cs`
- `FMReadiness_v3/UI/ExternalEvents/ParameterEditorExternalEventHandler.cs`
- `FMReadiness_v3/UI/app.js`
- `FMReadiness_v3/UI/index.html`
- `FMReadiness_v3/UI/styles.css`

## Troubleshooting
- Blank pane: install/repair WebView2 runtime.
- Audit shows no assets: verify model contains supported categories.
- Fix/selection confusion: check **Lock selection** state and refresh selection.
- IFC-sidecar mismatch: export IFC with stable GUID workflow and then export sidecar from the same model state.
