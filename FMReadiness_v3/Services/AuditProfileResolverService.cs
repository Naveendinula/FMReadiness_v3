using System.Collections.Generic;

namespace FMReadiness_v3.Services
{
    public static class AuditProfileState
    {
        private static readonly object SyncRoot = new();
        private static string? _activePresetFile;
        private static string? _activePresetName;

        public static void SetActivePreset(string? presetFile, string? presetName = null)
        {
            lock (SyncRoot)
            {
                _activePresetFile = string.IsNullOrWhiteSpace(presetFile) ? null : presetFile;
                _activePresetName = string.IsNullOrWhiteSpace(presetName) ? null : presetName;
            }
        }

        public static (string? presetFile, string? presetName) GetActivePreset()
        {
            lock (SyncRoot)
            {
                return (_activePresetFile, _activePresetName);
            }
        }
    }

    public class AuditProfileResolverService
    {
        public bool TryResolveRules(
            out Dictionary<string, CategoryConfig> rules,
            out string profileName,
            out string errorMessage)
        {
            var checklistService = new ChecklistService();
            var (presetFile, presetName) = AuditProfileState.GetActivePreset();

            if (!string.IsNullOrWhiteSpace(presetFile))
            {
                var activePresetFile = presetFile!;
                var presetService = new PresetService();
                if (presetService.LoadPreset(activePresetFile) && checklistService.LoadFromPreset(presetService))
                {
                    rules = checklistService.Rules;
                    profileName = !string.IsNullOrWhiteSpace(presetName)
                        ? presetName!
                        : checklistService.CurrentChecklistName;
                    errorMessage = string.Empty;
                    return true;
                }
            }

            if (checklistService.LoadConfig())
            {
                rules = checklistService.Rules;
                profileName = checklistService.CurrentChecklistName;
                errorMessage = string.Empty;
                return true;
            }

            rules = new Dictionary<string, CategoryConfig>();
            profileName = string.Empty;
            errorMessage = "Could not load the audit profile.";
            return false;
        }
    }
}
