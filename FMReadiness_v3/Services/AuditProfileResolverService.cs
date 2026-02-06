using System;
using System.Collections.Generic;

namespace FMReadiness_v3.Services
{
    public enum AuditScoreMode
    {
        RequiredOnly,
        AllEditable
    }

    public static class AuditProfileState
    {
        private static readonly object SyncRoot = new();
        private static string? _activePresetFile;
        private static string? _activePresetName;
        private static AuditScoreMode _scoreMode = AuditScoreMode.AllEditable;

        public static void SetActivePreset(string? presetFile, string? presetName = null)
        {
            lock (SyncRoot)
            {
                _activePresetFile = string.IsNullOrWhiteSpace(presetFile) ? null : presetFile;
                _activePresetName = string.IsNullOrWhiteSpace(presetName) ? null : presetName;
            }
        }

        public static void SetScoreMode(AuditScoreMode mode)
        {
            lock (SyncRoot)
            {
                _scoreMode = mode;
            }
        }

        public static void SetScoreMode(string? mode)
        {
            if (!TryParseScoreMode(mode, out var parsed))
                return;

            SetScoreMode(parsed);
        }

        public static AuditScoreMode GetScoreMode()
        {
            lock (SyncRoot)
            {
                return _scoreMode;
            }
        }

        public static string GetScoreModeKey(AuditScoreMode mode)
        {
            return mode == AuditScoreMode.RequiredOnly ? "required" : "all";
        }

        public static string GetScoreModeLabel(AuditScoreMode mode)
        {
            return mode == AuditScoreMode.RequiredOnly
                ? "Required + unique only"
                : "All editable fields";
        }

        private static bool TryParseScoreMode(string? mode, out AuditScoreMode scoreMode)
        {
            scoreMode = AuditScoreMode.AllEditable;
            if (string.IsNullOrWhiteSpace(mode))
                return false;

            if (string.Equals(mode, "required", StringComparison.OrdinalIgnoreCase)
                || string.Equals(mode, "requiredOnly", StringComparison.OrdinalIgnoreCase))
            {
                scoreMode = AuditScoreMode.RequiredOnly;
                return true;
            }

            if (string.Equals(mode, "all", StringComparison.OrdinalIgnoreCase)
                || string.Equals(mode, "allEditable", StringComparison.OrdinalIgnoreCase))
            {
                scoreMode = AuditScoreMode.AllEditable;
                return true;
            }

            return false;
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
