# FM Readiness v3

Revit add-in for auditing Facility Management (FM) data completeness, editing key FM parameters, and exporting FM metadata to a sidecar JSON for DigitalTwin/IFC workflows.

## What it does
- **Audit FM readiness** across supported categories and show results in a dockable UI.
- **Highlight missing instance/type data** and compute readiness scores per element and group.
- **Edit FM parameters in bulk** (selected elements or by category) directly in Revit.
- **Export a `.fm_params.json` sidecar** keyed by IFC GlobalId for DigitalTwin viewers.

## Supported Revit categories
The collector currently audits:
- `OST_MechanicalEquipment`
- `OST_DuctTerminal`
- `OST_DuctAccessory`
- `OST_PipeAccessory`

## FM parameters in scope
**Instance parameters**
- `FM_Barcode`
- `FM_UniqueAssetId`
- `FM_InstallationDate`
- `FM_WarrantyStart`
- `FM_WarrantyEnd`
- `FM_Criticality`
- `FM_Trade`
- `FM_PMTemplateId`
- `FM_PMFrequencyDays`
- `FM_Building`
- `FM_LocationSpace`

**Type parameters**
- `Manufacturer`
- `Model`
- `Type Mark` (Revit built-in `ALL_MODEL_TYPE_MARK`)

These are also defined in `FMReadiness_v3/IFC/FMReadiness_Psets.txt` for IFC property set mapping.

## Getting started

### Prerequisites
- Autodesk Revit installed (configs are provided for 2022–2026).
- Visual Studio (or Rider) with .NET desktop workload.
- WebView2 runtime (required for the dockable UI pane).

### Build
1. Open `FMReadiness_v3.slnx`.
2. Select a configuration that matches your Revit version (e.g., `Debug.R26` or `Release.R24`).
3. Build the solution.

> The project uses `Nice3point.Revit.Sdk` with `DeployAddin=true`. If your environment supports it, the build will deploy the add-in automatically to the Revit add-ins folder. If not, copy:
> - `FMReadiness_v3.addin`
> - `FMReadiness_v3.dll`
> into your Revit add-ins directory for the target version.

### Run in Revit
After launching Revit, you’ll find a **Digital Twin** tab with the **FM Tools** panel:
- **Run FM Audit** — runs the readiness audit and opens the results pane.
- **FM Pane** — show/hide the dockable pane.
- **Export FM Sidecar** — exports `*.fm_params.json` for DigitalTwin viewers.

## Audit checklist configuration
The audit rules are defined in `FMReadiness_v3/readiness_checklist.json`. The file maps categories to groups and fields:

```json
{
  "OST_MechanicalEquipment": {
    "groups": {
      "Identity": {
        "fields": [
          {
            "key": "AssetTag",
            "label": "Asset Tag",
            "scope": "instance",
            "source": { "type": "name", "value": "Asset Tag" },
            "rules": ["unique"]
          }
        ]
      }
    }
  }
}
```

**Field source types**
- `name` — parameter by display name.
- `builtin` — Revit built-in parameter (e.g., `ALL_MODEL_TYPE_MARK`).
- `sharedGuid` — shared parameter by GUID.
- `computed` — computed value:
  - `Element.UniqueId`
  - `Element.LevelName`
  - `Element.RoomOrSpace`

**Rules**
- `unique` — flags duplicates (e.g., asset IDs).
- `date` — used for UI display/validation.

> Edit the JSON, rebuild, and the updated file will be copied to the add-in output.

## Sidecar export (.fm_params.json)
The export is keyed by **IFC GlobalId**. If Revit has stored IFC GUIDs, those are used; otherwise a computed fallback is used.

Example structure:
```json
{
  "3kT$VfZq1A1QYF8G3E4X2a": {
    "FMReadiness": {
      "FM_Barcode": "MECH-001",
      "FM_Criticality": "3"
    },
    "FMReadinessType": {
      "Manufacturer": "ACME",
      "Model": "X100"
    },
    "_meta": {
      "RevitElementId": 123456,
      "RevitUniqueId": "....",
      "Category": "Mechanical Equipment",
      "Family": "AHU",
      "TypeName": "AHU-1"
    }
  }
}
```

**Recommended workflow**
1. Export IFC with **Store IFC GUID** enabled (File → Export → IFC → Modify Setup → General tab).
2. Export FM sidecar from the add-in.
3. Upload both the `.ifc` and `.fm_params.json` to the DigitalTwin viewer.

## UI overview
The dockable pane (WebView2) provides:
- **Audit Results**: summary metrics, group scores, searchable table of missing data.
- **Parameter Editor**: update FM instance parameters for selected elements.
- **Category Bulk Fill**: apply values across a category (optional fill-only-blank).
- **Type Parameters**: edit Manufacturer/Model/Type Mark.
- **Computed values**: UniqueId, Level, Room/Space helpers.

UI assets live in `FMReadiness_v3/UI/` and are copied to the build output.

## Project layout
- `FMReadiness_v3/Commands/` — Revit ribbon commands (audit, pane toggle, sidecar export).
- `FMReadiness_v3/Services/` — audit logic, checklist loader, sidecar exporter.
- `FMReadiness_v3/UI/` — WebView2 UI assets.
- `FMReadiness_v3/IFC/FMReadiness_Psets.txt` — IFC property set mapping.
- `FMReadiness_v3/readiness_checklist.json` — configurable audit rules.

## Notes & troubleshooting
- If the pane is blank, verify WebView2 runtime is installed.
- If the audit finds nothing, confirm the model contains the supported categories.
- If the sidecar shows mismatched elements in the viewer, export IFC with **Store IFC GUID** enabled.
