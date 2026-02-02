# FM Readiness v3

Revit add-in for auditing Facility Management (FM) data completeness, editing key FM parameters, and exporting FM metadata to a sidecar JSON for DigitalTwin/IFC workflows. Supports **COBie 2.4** field mapping with flexible preset configurations.

## What it does
- **Audit FM readiness** across supported categories and show results in a dockable UI.
- **Highlight missing instance/type data** and compute readiness scores per element and group.
- **Edit FM parameters in bulk** (selected elements or by category) directly in Revit.
- **Export a `.fm_params.json` sidecar** keyed by IFC GlobalId for DigitalTwin viewers.
- **COBie 2.4 support** with presets for field mapping, aliases, and validation.

## Supported Revit categories
The collector currently audits:
- `OST_MechanicalEquipment`
- `OST_DuctTerminal`
- `OST_DuctAccessory`
- `OST_PipeAccessory`
- `OST_Rooms` (for Space data)
- `OST_MEPSpaces` (for Space data)
- `OST_Levels` (for Floor data)

## COBie 2.4 Support

### Preset System
FM Readiness v3 now includes a **preset system** for mapping Revit parameters to COBie 2.4 fields. Presets are JSON files located in the `Presets/` folder.

**Built-in presets:**
- `cobie-core.json` — Core COBie 2.4 fields (Component, Type, Space, Floor, Facility tables)
- `fm-legacy.json` — Maps existing FM_* parameters to COBie equivalents
- `custom.json` — User-editable template for custom mappings

### COBie Tables Supported
| COBie Table | Revit Categories |
|-------------|------------------|
| Component | MechanicalEquipment, DuctTerminal, DuctAccessory, PipeAccessory |
| Type | Element Types from above categories |
| Space | Rooms, MEP Spaces |
| Floor | Levels |
| Facility | Project Information |

### Field Mapping with Aliases
Each COBie field can specify a primary Revit parameter and fallback aliases. The system reads from the first parameter that has a value:

```json
{
  "cobieKey": "Component.SerialNumber",
  "label": "Serial Number",
  "revitParam": "COBie_SerialNumber",
  "aliasParams": ["FM_Barcode", "Equipment Serial Number", "Asset Tag"],
  "rules": ["unique"]
}
```

**Read behavior:** Tries `COBie_SerialNumber` → `FM_Barcode` → `Equipment Serial Number` → `Asset Tag`

**Write behavior:** Always writes to the primary parameter (`COBie_SerialNumber`)

### Computed Values
Some fields support computed values from Revit metadata:
- `Element.UniqueId` — Revit's unique element identifier
- `Element.LevelName` — Level name from element placement
- `Element.RoomOrSpace` — Room/Space name from element location
- `Type.Manufacturer` — From type parameter
- `Type.Model` — From type parameter

### Validation Rules
- `required` — Field must have a value (highlighted in audit)
- `unique` — Value must be unique across all elements (duplicates flagged)
- `date` — Value should be in ISO 8601 format (YYYY-MM-DD)

## FM parameters in scope
**Instance parameters (COBie Component)**
- `COBie_SerialNumber` / `FM_Barcode` — Asset Tag/Serial Number
- `COBie_InstallationDate` / `FM_InstallationDate` — Installation Date
- `COBie_WarrantyStartDate` / `FM_WarrantyStart` — Warranty Start
- `COBie_WarrantyDurationLabor` — Warranty Labor Duration
- `COBie_WarrantyDurationParts` — Warranty Parts Duration
- `FM_UniqueAssetId` — Unique Asset Identifier
- `FM_Criticality` — Criticality Rating
- `FM_Trade` — Trade/Discipline

**Type parameters (COBie Type)**
- `Manufacturer` — Equipment manufacturer
- `Model` — Model number
- `Type Mark` — Type identifier
- `COBie_ModelNumber` — COBie model reference
- `COBie_AssetType` — Asset classification

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
The audit rules are defined in `FMReadiness_v3/readiness_checklist.json` (legacy format) or loaded from preset files in `Presets/` folder. The file maps categories to groups and fields:

### Legacy Format (readiness_checklist.json)
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

### COBie Preset Format (Presets/*.json)
```json
{
  "name": "COBie Core",
  "description": "Core COBie 2.4 fields",
  "version": "1.0.0",
  "tables": {
    "Component": {
      "scope": "instance",
      "categories": ["OST_MechanicalEquipment", "OST_DuctTerminal"],
      "fields": [
        {
          "cobieKey": "Component.SerialNumber",
          "label": "Serial Number",
          "scope": "instance",
          "dataType": "string",
          "required": true,
          "revitParam": "COBie_SerialNumber",
          "aliasParams": ["FM_Barcode", "Asset Tag"],
          "rules": ["unique"],
          "group": "Identity"
        }
      ]
    }
  }
}
```

**Preset Field Properties:**
| Property | Description |
|----------|-------------|
| `cobieKey` | COBie table and field reference (e.g., `Component.SerialNumber`) |
| `label` | Display name in UI |
| `scope` | `instance` or `type` |
| `dataType` | `string`, `date`, `number`, `boolean` |
| `required` | Whether field is required for COBie compliance |
| `revitParam` | Primary Revit parameter name |
| `aliasParams` | Fallback parameter names (array) |
| `rules` | Validation rules: `unique`, `date`, `required` |
| `group` | Grouping for UI display |
| `computedFrom` | For computed values: `{ "property": "Element.UniqueId" }` |

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
- `required` — flags missing values as validation errors.

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
- `FMReadiness_v3/Services/` — audit logic, checklist loader, COBie mapping, sidecar exporter.
  - `AuditService.cs` — Core audit logic with scoring and validation.
  - `ChecklistService.cs` — Loads checklist configurations.
  - `CollectorService.cs` — Collects Revit elements by category.
  - `PresetService.cs` — Loads and manages COBie presets.
  - `CobieMappingService.cs` — COBie field resolution with alias support.
- `FMReadiness_v3/UI/` — WebView2 UI assets.
- `FMReadiness_v3/Presets/` — COBie preset configuration files.
  - `cobie-core.json` — Core COBie 2.4 field definitions.
  - `fm-legacy.json` — FM_* parameter to COBie mappings.
  - `custom.json` — User-editable template.
- `FMReadiness_v3/IFC/FMReadiness_Psets.txt` — IFC property set mapping.
- `FMReadiness_v3/readiness_checklist.json` — Legacy configurable audit rules.
- `FMReadiness_v3/cobie-core-checklist.json` — COBie-based audit checklist.

## Migration Guide

### Migrating from FM_* parameters to COBie
If you have existing FM data using the FM_* parameter naming convention, the `fm-legacy.json` preset provides backwards compatibility:

1. Select the **FM Legacy** preset from the dropdown in the UI.
2. The audit will read values from your existing FM_* parameters.
3. When editing, values are written to the COBie-named parameters.
4. Optionally, use the "Copy to COBie" feature to migrate values.

### Creating Custom Presets
1. Copy `Presets/custom.json` to a new file (e.g., `my-company.json`).
2. Edit the field mappings to match your parameter naming.
3. Rebuild or copy to the output folder.
4. Select your preset from the UI dropdown.

### Parameter Naming Recommendations
For new projects, use COBie-named parameters directly:
- `COBie_SerialNumber` instead of `FM_Barcode`
- `COBie_InstallationDate` instead of `FM_InstallationDate`
- `COBie_WarrantyStartDate` instead of `FM_WarrantyStart`

This ensures compatibility with standard COBie export tools.

## Notes & troubleshooting
- If the pane is blank, verify WebView2 runtime is installed.
- If the audit finds nothing, confirm the model contains the supported categories.
- If the sidecar shows mismatched elements in the viewer, export IFC with **Store IFC GUID** enabled.
