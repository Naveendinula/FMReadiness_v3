using System;
using System.IO;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using FMReadiness_v3.Services;
using Microsoft.Win32;
using Nice3point.Revit.Toolkit.External;

namespace FMReadiness_v3.Commands
{
    /// <summary>
    /// Exports FM parameters as a sidecar JSON file that can be merged with
    /// the DigitalTwin viewer's metadata.json.
    /// 
    /// The sidecar file is keyed by IFC GlobalId, allowing FM parameters to
    /// appear in the viewer's Property Panel as "FMReadiness" and "FMReadinessType"
    /// property sets.
    /// </summary>
    [UsedImplicitly]
    [Transaction(TransactionMode.ReadOnly)]
    public class ExportFmSidecarCommand : ExternalCommand
    {
        public override void Execute()
        {
            var doc = Context.ActiveDocument ?? Context.Document;
            if (doc == null)
            {
                TaskDialog.Show("FM Readiness", "No active document. Open a model and try again.");
                return;
            }

            // Prompt for output path
            var outputPath = PromptForOutputPath(doc);
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                return; // User cancelled
            }

            try
            {
                var service = new FmSidecarExportService(doc);
                var result = service.Export(outputPath);

                // Show results
                var message = new StringBuilder();
                message.AppendLine("FM Sidecar Export Complete");
                message.AppendLine();
                message.AppendLine($"File: {result.OutputPath}");
                message.AppendLine($"Size: {result.FileSizeBytes / 1024.0:F1} KB");
                message.AppendLine();
                message.AppendLine($"Total FM elements: {result.TotalElements}");
                message.AppendLine($"Exported: {result.ExportedElements}");

                // Show IFC GUID source breakdown
                if (result.StoredGuidCount > 0 || result.ComputedGuidCount > 0)
                {
                    message.AppendLine();
                    message.AppendLine("IFC GUID sources:");
                    if (result.StoredGuidCount > 0)
                        message.AppendLine($"  - Stored (from IFC export): {result.StoredGuidCount}");
                    if (result.ComputedGuidCount > 0)
                        message.AppendLine($"  - Computed (fallback): {result.ComputedGuidCount}");
                }

                message.AppendLine();
                message.AppendLine($"Skipped (no data): {result.SkippedNoData}");
                message.AppendLine($"Skipped (no GUID): {result.SkippedNoGuid}");

                if (result.Errors.Count > 0)
                {
                    message.AppendLine();
                    message.AppendLine($"Errors: {result.Errors.Count}");
                    foreach (var error in result.Errors.GetRange(0, Math.Min(5, result.Errors.Count)))
                    {
                        message.AppendLine($"  - {error}");
                    }
                    if (result.Errors.Count > 5)
                    {
                        message.AppendLine($"  ... and {result.Errors.Count - 5} more");
                    }
                }

                message.AppendLine();
                message.AppendLine("Next steps:");
                message.AppendLine("1. Export the model to IFC with 'Store IFC GUID' enabled");
                message.AppendLine("   (File > Export > IFC > Modify Setup > General tab)");
                message.AppendLine("2. Upload both files to DigitalTwin viewer:");
                message.AppendLine("   - The .ifc file");
                message.AppendLine($"   - The {Path.GetFileName(outputPath)} sidecar file");
                message.AppendLine("3. FM parameters will appear in the Property Panel");

                if (result.ComputedGuidCount > 0 && result.StoredGuidCount == 0)
                {
                    message.AppendLine();
                    message.AppendLine("TIP: Enable 'Store IFC GUID' during IFC export for");
                    message.AppendLine("more reliable element matching in the viewer.");
                }

                TaskDialog.Show("FM Readiness", message.ToString());
            }
            catch (Exception ex)
            {
                TaskDialog.Show("FM Readiness - Export Failed", $"Error: {ex.Message}");
            }
        }

        private static string? PromptForOutputPath(Document doc)
        {
            var defaultName = GetSafeFileName($"{doc.Title}_FM.fm_params.json");

            var dialog = new SaveFileDialog
            {
                Title = "Export FM Sidecar (for DigitalTwin Viewer)",
                Filter = "FM Parameters JSON (*.fm_params.json)|*.fm_params.json|All JSON files (*.json)|*.json",
                DefaultExt = ".fm_params.json",
                FileName = defaultName,
                OverwritePrompt = true
            };

            var result = dialog.ShowDialog();
            if (result != true) return null;

            return dialog.FileName;
        }

        private static string GetSafeFileName(string name)
        {
            foreach (var invalid in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(invalid, '_');
            }
            return name;
        }
    }
}
