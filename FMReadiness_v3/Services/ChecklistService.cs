using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using Autodesk.Revit.UI;

namespace FMReadiness_v3.Services
{
    public class ChecklistService
    {
        private const string DefaultChecklist = "readiness_checklist.json";
        private const string CobieChecklist = "cobie-core-checklist.json";

        public Dictionary<string, CategoryConfig> Rules { get; private set; } = new();
        public string CurrentChecklistName { get; private set; } = string.Empty;

        private readonly string _assemblyDir;

        public ChecklistService()
        {
            var assemblyPath = Assembly.GetExecutingAssembly().Location;
            _assemblyDir = Path.GetDirectoryName(assemblyPath) ?? string.Empty;
        }

        /// <summary>
        /// Loads the default checklist configuration.
        /// </summary>
        public bool LoadConfig()
        {
            return LoadChecklist(DefaultChecklist);
        }

        /// <summary>
        /// Loads a specific checklist file.
        /// </summary>
        public bool LoadChecklist(string fileName)
        {
            try
            {
                var configPath = Path.Combine(_assemblyDir, fileName);

                if (!File.Exists(configPath))
                {
                    // Try in Presets folder
                    configPath = Path.Combine(_assemblyDir, "Presets", fileName);
                }

                if (!File.Exists(configPath))
                {
                    TaskDialog.Show("FM Readiness", $"Checklist file not found: {fileName}");
                    return false;
                }

                var jsonContent = File.ReadAllText(configPath);
                var settings = new DataContractJsonSerializerSettings
                {
                    UseSimpleDictionaryFormat = true
                };
                var serializer = new DataContractJsonSerializer(typeof(Dictionary<string, CategoryConfig>), settings);
                using var stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonContent));
                var rules = serializer.ReadObject(stream) as Dictionary<string, CategoryConfig>;
                if (rules is null) return false;

                Rules = rules;
                CurrentChecklistName = fileName;
                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("FM Readiness", $"Failed to load checklist JSON: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Loads the COBie core checklist.
        /// </summary>
        public bool LoadCobieChecklist()
        {
            return LoadChecklist(CobieChecklist);
        }

        /// <summary>
        /// Loads rules from a PresetService configuration.
        /// </summary>
        public bool LoadFromPreset(PresetService presetService)
        {
            if (presetService?.CurrentPreset == null)
                return false;

            try
            {
                Rules = presetService.ConvertToChecklistRules();
                CurrentChecklistName = $"Preset: {presetService.CurrentPreset.Name}";
                return Rules.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Gets available checklist files.
        /// </summary>
        public List<string> GetAvailableChecklists()
        {
            var checklists = new List<string>();

            // Check root directory
            foreach (var file in Directory.GetFiles(_assemblyDir, "*checklist*.json"))
            {
                checklists.Add(Path.GetFileName(file));
            }

            // Check Presets folder
            var presetsPath = Path.Combine(_assemblyDir, "Presets");
            if (Directory.Exists(presetsPath))
            {
                foreach (var file in Directory.GetFiles(presetsPath, "*checklist*.json"))
                {
                    checklists.Add(Path.GetFileName(file));
                }
            }

            return checklists;
        }
    }

    [DataContract]
    public class CategoryConfig
    {
        [DataMember(Name = "groups")]
        public Dictionary<string, GroupConfig> Groups { get; set; } = new();
    }

    [DataContract]
    public class GroupConfig
    {
        [DataMember(Name = "fields")]
        public List<FieldSpec> Fields { get; set; } = new();
    }

    [DataContract]
    public class FieldSpec
    {
        [DataMember(Name = "key")]
        public string Key { get; set; } = string.Empty;
        [DataMember(Name = "label")]
        public string Label { get; set; } = string.Empty;
        [DataMember(Name = "scope")]
        public string Scope { get; set; } = "instance"; // instance | type | either
        [DataMember(Name = "source")]
        public FieldSource Source { get; set; } = new();
        [DataMember(Name = "required")]
        public bool? Required { get; set; }
        [DataMember(Name = "rules")]
        public List<string> Rules { get; set; } = new();
    }

    [DataContract]
    public class FieldSource
    {
        [DataMember(Name = "type")]
        public string Type { get; set; } = "name"; // builtin | name | sharedGuid | computed
        [DataMember(Name = "value")]
        public string? Value { get; set; }
        [DataMember(Name = "id")]
        public string? Id { get; set; }
    }
}
