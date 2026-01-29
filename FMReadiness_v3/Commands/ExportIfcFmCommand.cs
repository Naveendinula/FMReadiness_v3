using System;
using System.IO;
using System.Text;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using FMReadiness_v3.IFC;
using Microsoft.Win32;
using Nice3point.Revit.Toolkit.External;

namespace FMReadiness_v3.Commands
{
    [UsedImplicitly]
    [Transaction(TransactionMode.Manual)]
    public class ExportIfcFmCommand : ExternalCommand
    {
        public override void Execute()
        {
            var doc = Context.ActiveDocument ?? Context.Document;
            if (doc == null)
            {
                TaskDialog.Show("FM Readiness", "No active document. Open a model and try again.");
                return;
            }

            var psetPath = IfcExportHelper.EnsureUserDefinedPsetFile();
            if (string.IsNullOrWhiteSpace(psetPath))
            {
                TaskDialog.Show("FM Readiness", "Unable to create the IFC property set configuration file.");
                return;
            }

            if (!TryGetIncludeRevitPropertySets(out var includeRevitPropertySets))
            {
                return;
            }

            var exportPath = PromptForIfcFilePath(doc);
            if (string.IsNullOrWhiteSpace(exportPath))
            {
                return;
            }

            var exportFolder = Path.GetDirectoryName(exportPath);
            if (string.IsNullOrWhiteSpace(exportFolder))
            {
                TaskDialog.Show("FM Readiness", "Invalid export folder.");
                return;
            }

            var exportFileName = Path.GetFileName(exportPath);
            var options = new IFCExportOptions();
            IfcExportHelper.ConfigureIfc4Options(options, psetPath, includeRevitPropertySets);

            if (doc.IsReadOnly)
            {
                TaskDialog.Show("FM Readiness", "Document is read-only. Make it editable before exporting.");
                return;
            }

            Transaction? tx = null;
            try
            {
                WriteExportLog(exportPath, "Starting IFC export", log =>
                {
                    log.AppendLine($"Document: {doc.Title}");
                    log.AppendLine($"Document Path: {doc.PathName}");
                    log.AppendLine($"Export Path: {exportPath}");
                    log.AppendLine($"Pset File: {psetPath}");
                    log.AppendLine($"Include Revit Property Sets: {includeRevitPropertySets}");
                    log.AppendLine("Options:");
                    log.AppendLine("  FileVersion = IFC4DTV");
                    log.AppendLine($"  ExportBaseQuantities = {includeRevitPropertySets}");
                    log.AppendLine("  ExportUserDefinedPsets = true");
                    log.AppendLine($"  ExportUserDefinedPsetsFileName = {psetPath}");
                    log.AppendLine("  ExportUserDefinedParameterMapping = false");
                    log.AppendLine($"  ExportInternalRevitPropertySets = {includeRevitPropertySets}");
                    log.AppendLine($"  ExportIFCCommonPropertySets = {includeRevitPropertySets}");
                    log.AppendLine($"  ExportMaterialPsets = {includeRevitPropertySets}");
                    log.AppendLine("  StoreIFCGUID = false");
                    log.AppendLine("  FileType = IFC");
                    log.AppendLine($"  OptionProperties: {IfcExportHelper.GetOptionPropertyDiagnostics(options)}");
                });

                tx = new Transaction(doc, "Export IFC4 (FM)");
                if (tx.Start() != TransactionStatus.Started)
                {
                    TaskDialog.Show("FM Readiness", "Could not start a transaction for IFC export.");
                    return;
                }

                var ok = doc.Export(exportFolder, exportFileName, options);
                if (ok)
                {
                    tx.Commit();
                    WriteExportLog(exportPath, "IFC export completed", log =>
                    {
                        LogExportFileDiagnostics(log, exportPath);
                    });
                    TaskDialog.Show("FM Readiness", $"IFC export complete:\n{exportPath}");
                }
                else
                {
                    tx.RollBack();
                    TaskDialog.Show("FM Readiness - IFC Export Failed", "IFC export returned false.");
                }
            }
            catch (Exception ex)
            {
                if (tx != null && tx.GetStatus() == TransactionStatus.Started)
                {
                    try { tx.RollBack(); } catch { }
                }

                WriteExportLog(exportPath, "IFC export failed", log =>
                {
                    log.AppendLine($"Exception: {ex}");
                });
                TaskDialog.Show("FM Readiness - IFC Export Failed", ex.Message);
            }
        }

        private static string? PromptForIfcFilePath(Document doc)
        {
            var defaultName = GetSafeFileName($"{doc.Title}_FM.ifc");

            var dialog = new SaveFileDialog
            {
                Title = "Export IFC (FM Readiness)",
                Filter = "IFC files (*.ifc)|*.ifc",
                DefaultExt = ".ifc",
                FileName = defaultName,
                OverwritePrompt = true
            };

            var result = dialog.ShowDialog();
            if (result != true) return null;

            return dialog.FileName;
        }

        private static bool TryGetIncludeRevitPropertySets(out bool includeRevitPropertySets)
        {
            includeRevitPropertySets = false;

            var dialog = new TaskDialog("FM Readiness - IFC Export")
            {
                MainInstruction = "Include Revit property sets?",
                MainContent =
                    "FM user-defined property sets will always be included.\n" +
                    "Choose Yes to also include Revit property sets, or No for FM-only export.",
                CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No | TaskDialogCommonButtons.Cancel
            };

            var result = dialog.Show();
            if (result == TaskDialogResult.Cancel)
            {
                return false;
            }

            includeRevitPropertySets = result == TaskDialogResult.Yes;
            return true;
        }

        private static string GetSafeFileName(string name)
        {
            foreach (var invalid in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(invalid, '_');
            }

            return name;
        }

        private static void WriteExportLog(string exportPath, string title, Action<StringBuilder> writer)
        {
            try
            {
                var logPath = GetExportLogPath(exportPath);

                var sb = new StringBuilder();
                sb.AppendLine($"[{DateTime.UtcNow:O}] {title}");
                writer(sb);
                sb.AppendLine();

                File.AppendAllText(logPath, sb.ToString());
            }
            catch
            {
                TryWriteFallbackLog(title, writer);
            }
        }

        private static string GetExportLogPath(string exportPath)
        {
            var exportFolder = Path.GetDirectoryName(exportPath);
            var fileName = Path.GetFileNameWithoutExtension(exportPath);
            if (string.IsNullOrWhiteSpace(exportFolder))
            {
                exportFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            }

            return Path.Combine(exportFolder, $"{fileName}_FMExport.log");
        }

        private static void TryWriteFallbackLog(string title, Action<StringBuilder> writer)
        {
            try
            {
                var logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "FMReadiness_v3",
                    "logs");
                Directory.CreateDirectory(logDir);
                var logPath = Path.Combine(logDir, "ifc-export.log");

                var sb = new StringBuilder();
                sb.AppendLine($"[{DateTime.UtcNow:O}] {title}");
                writer(sb);
                sb.AppendLine();

                File.AppendAllText(logPath, sb.ToString());
            }
            catch
            {
                // Ignore logging failures.
            }
        }

        private static void LogExportFileDiagnostics(StringBuilder log, string exportPath)
        {
            if (!File.Exists(exportPath))
            {
                log.AppendLine("Export file not found after export.");
                return;
            }

            var fileInfo = new FileInfo(exportPath);
            log.AppendLine($"Export file size: {fileInfo.Length} bytes");

            using (var fs = new FileStream(exportPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var header = new byte[4];
                var read = fs.Read(header, 0, header.Length);
                var headerText = Encoding.ASCII.GetString(header, 0, read);
                log.AppendLine($"File header: {headerText.Replace("\0", string.Empty)}");

                if (headerText.StartsWith("PK", StringComparison.OrdinalIgnoreCase))
                {
                    log.AppendLine("File appears to be ZIP (IFCZIP). Cannot scan for Pset strings.");
                    return;
                }
            }

            var fmCount = CountOccurrences(exportPath, "FMReadiness");
            var revitCount = CountOccurrences(exportPath, "Pset_Revit");
            log.AppendLine($"Occurrences: FMReadiness={fmCount}, Pset_Revit={revitCount}");
        }

        private static int CountOccurrences(string path, string token)
        {
            try
            {
                var count = 0;
                using var reader = new StreamReader(path);
                string? line;
                while ((line = reader.ReadLine()) != null)
                {
                    var index = 0;
                    while (true)
                    {
                        index = line.IndexOf(token, index, StringComparison.OrdinalIgnoreCase);
                        if (index < 0) break;
                        count++;
                        index += token.Length;
                    }
                }

                return count;
            }
            catch
            {
                return -1;
            }
        }
    }
}
