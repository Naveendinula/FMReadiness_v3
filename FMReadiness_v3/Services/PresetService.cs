using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace FMReadiness_v3.Services
{
    /// <summary>
    /// Service for loading and managing COBie/FM preset configurations.
    /// </summary>
    public class PresetService
    {
        private const string PresetsFolder = "Presets";
        private const string DefaultPreset = "cobie-core.json";

        public CobiePreset? CurrentPreset { get; private set; }
        public string CurrentPresetName { get; private set; } = string.Empty;

        private readonly string _presetsPath;

        public PresetService()
        {
            var assemblyPath = Assembly.GetExecutingAssembly().Location;
            var assemblyDir = Path.GetDirectoryName(assemblyPath) ?? string.Empty;
            _presetsPath = Path.Combine(assemblyDir, PresetsFolder);
        }

        /// <summary>
        /// Gets available preset files.
        /// </summary>
        public List<PresetInfo> GetAvailablePresets()
        {
            var presets = new List<PresetInfo>();

            if (!Directory.Exists(_presetsPath))
                return presets;

            foreach (var file in Directory.GetFiles(_presetsPath, "*.json"))
            {
                try
                {
                    var fileName = Path.GetFileName(file);
                    if (string.Equals(fileName, "custom.json", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var preset = LoadPresetFile(file);
                    if (preset != null)
                    {
                        presets.Add(new PresetInfo
                        {
                            FileName = fileName,
                            Name = preset.Name ?? Path.GetFileNameWithoutExtension(file),
                            Description = preset.Description ?? string.Empty,
                            Version = preset.Version ?? "1.0.0"
                        });
                    }
                }
                catch
                {
                    // Skip invalid preset files
                }
            }

            return presets;
        }

        /// <summary>
        /// Loads a preset by filename.
        /// </summary>
        public bool LoadPreset(string fileName)
        {
            var filePath = Path.Combine(_presetsPath, fileName);
            if (!File.Exists(filePath))
                return false;

            var preset = LoadPresetFile(filePath);
            if (preset == null)
                return false;

            CurrentPreset = preset;
            CurrentPresetName = fileName;
            return true;
        }

        /// <summary>
        /// Loads the default COBie core preset.
        /// </summary>
        public bool LoadDefaultPreset()
        {
            return LoadPreset(DefaultPreset);
        }

        /// <summary>
        /// Gets all fields from the current preset organized by group.
        /// </summary>
        public Dictionary<string, List<CobieFieldSpec>> GetFieldsByGroup(string tableName = "Component")
        {
            var result = new Dictionary<string, List<CobieFieldSpec>>();
            if (CurrentPreset?.Tables == null)
                return result;

            if (!CurrentPreset.Tables.TryGetValue(tableName, out var table) || table?.Fields == null)
                return result;

            foreach (var field in table.Fields)
            {
                if (field == null) continue;
                var group = field.Group ?? "Other";

                if (!result.ContainsKey(group))
                    result[group] = new List<CobieFieldSpec>();

                result[group].Add(field);
            }

            // Add FM Ops extensions
            if (CurrentPreset.FmOpsExtensions?.Fields != null)
            {
                foreach (var field in CurrentPreset.FmOpsExtensions.Fields)
                {
                    if (field == null) continue;
                    var group = field.Group ?? "FMOps";

                    if (!result.ContainsKey(group))
                        result[group] = new List<CobieFieldSpec>();

                    result[group].Add(field);
                }
            }

            return result;
        }

        /// <summary>
        /// Gets all fields from the current preset as a flat list.
        /// </summary>
        public List<CobieFieldSpec> GetAllFields(string tableName = "Component")
        {
            var fields = new List<CobieFieldSpec>();
            if (CurrentPreset?.Tables == null)
                return fields;

            if (CurrentPreset.Tables.TryGetValue(tableName, out var table) && table?.Fields != null)
                fields.AddRange(table.Fields.Where(f => f != null));

            if (CurrentPreset.FmOpsExtensions?.Fields != null)
                fields.AddRange(CurrentPreset.FmOpsExtensions.Fields.Where(f => f != null));

            return fields;
        }

        /// <summary>
        /// Converts preset fields to ChecklistService-compatible rules.
        /// </summary>
        public Dictionary<string, CategoryConfig> ConvertToChecklistRules()
        {
            var rules = new Dictionary<string, CategoryConfig>();
            if (CurrentPreset?.Categories == null)
                return rules;

            var fieldsByGroup = GetFieldsByGroup("Component");
            var typeFields = GetAllFields("Type");

            foreach (var category in CurrentPreset.Categories)
            {
                var config = new CategoryConfig
                {
                    Groups = new Dictionary<string, GroupConfig>()
                };

                foreach (var groupEntry in fieldsByGroup)
                {
                    var groupName = groupEntry.Key;
                    var fields = groupEntry.Value;

                    var groupConfig = new GroupConfig
                    {
                        Fields = fields.Select(f => ConvertToFieldSpec(f)).ToList()
                    };

                    config.Groups[groupName] = groupConfig;
                }

                // Add type fields to MakeModel group if not already present
                if (!config.Groups.ContainsKey("MakeModel"))
                {
                    config.Groups["MakeModel"] = new GroupConfig { Fields = new List<FieldSpec>() };
                }

                foreach (var typeField in typeFields.Where(f => f.Scope == "type"))
                {
                    var existingKeys = config.Groups["MakeModel"].Fields.Select(f => f.Key).ToHashSet();
                    var cobieKey = typeField.CobieKey ?? string.Empty;
                    if (!string.IsNullOrEmpty(cobieKey) && !existingKeys.Contains(cobieKey))
                    {
                        config.Groups["MakeModel"].Fields.Add(ConvertToFieldSpec(typeField));
                    }
                }

                rules[category] = config;
            }

            return rules;
        }

        private FieldSpec ConvertToFieldSpec(CobieFieldSpec cobieField)
        {
            var spec = new FieldSpec
            {
                Key = cobieField.CobieKey ?? string.Empty,
                Label = cobieField.Label ?? string.Empty,
                Scope = cobieField.Scope ?? "instance",
                Rules = cobieField.Rules ?? new List<string>(),
                Source = new FieldSource()
            };

            // Determine source type
            if (cobieField.Computed != null && !string.IsNullOrEmpty(cobieField.Computed.Source))
            {
                spec.Source.Type = "computed";
                spec.Source.Id = cobieField.Computed.Source;
            }
            else if (!string.IsNullOrEmpty(cobieField.RevitBuiltIn))
            {
                spec.Source.Type = "builtin";
                spec.Source.Id = cobieField.RevitBuiltIn;
            }
            else if (!string.IsNullOrEmpty(cobieField.RevitParam))
            {
                spec.Source.Type = "name";
                spec.Source.Value = cobieField.RevitParam;
            }
            else if (cobieField.AliasParams != null && cobieField.AliasParams.Count > 0)
            {
                spec.Source.Type = "name";
                spec.Source.Value = cobieField.AliasParams[0];
            }

            return spec;
        }

        private CobiePreset? LoadPresetFile(string filePath)
        {
            try
            {
                var jsonContent = File.ReadAllText(filePath);
                var settings = new DataContractJsonSerializerSettings
                {
                    UseSimpleDictionaryFormat = true
                };
                var serializer = new DataContractJsonSerializer(typeof(CobiePreset), settings);
                using var stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonContent));
                return serializer.ReadObject(stream) as CobiePreset;
            }
            catch
            {
                return null;
            }
        }
    }

    #region Preset Data Contracts

    [DataContract]
    public class PresetInfo
    {
        [DataMember(Name = "fileName")]
        public string FileName { get; set; } = string.Empty;

        [DataMember(Name = "name")]
        public string Name { get; set; } = string.Empty;

        [DataMember(Name = "description")]
        public string Description { get; set; } = string.Empty;

        [DataMember(Name = "version")]
        public string Version { get; set; } = string.Empty;
    }

    [DataContract]
    public class CobiePreset
    {
        [DataMember(Name = "name")]
        public string? Name { get; set; }

        [DataMember(Name = "version")]
        public string? Version { get; set; }

        [DataMember(Name = "description")]
        public string? Description { get; set; }

        [DataMember(Name = "writeAliases")]
        public bool WriteAliases { get; set; } = true;

        [DataMember(Name = "categories")]
        public List<string>? Categories { get; set; }

        [DataMember(Name = "tables")]
        public Dictionary<string, CobieTable>? Tables { get; set; }

        [DataMember(Name = "fmOpsExtensions")]
        public CobieExtension? FmOpsExtensions { get; set; }

        [DataMember(Name = "computedFields")]
        public List<CobieFieldSpec>? ComputedFields { get; set; }
    }

    [DataContract]
    public class CobieTable
    {
        [DataMember(Name = "description")]
        public string? Description { get; set; }

        [DataMember(Name = "fields")]
        public List<CobieFieldSpec>? Fields { get; set; }
    }

    [DataContract]
    public class CobieExtension
    {
        [DataMember(Name = "description")]
        public string? Description { get; set; }

        [DataMember(Name = "fields")]
        public List<CobieFieldSpec>? Fields { get; set; }
    }

    [DataContract]
    public class CobieFieldSpec
    {
        [DataMember(Name = "cobieKey")]
        public string? CobieKey { get; set; }

        [DataMember(Name = "label")]
        public string? Label { get; set; }

        [DataMember(Name = "scope")]
        public string? Scope { get; set; } = "instance";

        [DataMember(Name = "dataType")]
        public string? DataType { get; set; } = "string";

        [DataMember(Name = "required")]
        public bool Required { get; set; }

        [DataMember(Name = "revitParam")]
        public string? RevitParam { get; set; }

        [DataMember(Name = "revitBuiltIn")]
        public string? RevitBuiltIn { get; set; }

        [DataMember(Name = "aliasParams")]
        public List<string>? AliasParams { get; set; }

        [DataMember(Name = "rules")]
        public List<string>? Rules { get; set; }

        [DataMember(Name = "group")]
        public string? Group { get; set; }

        [DataMember(Name = "cobieColumn")]
        public string? CobieColumn { get; set; }

        [DataMember(Name = "defaultValue")]
        public string? DefaultValue { get; set; }

        [DataMember(Name = "computed")]
        public CobieComputedSource? Computed { get; set; }
    }

    [DataContract]
    public class CobieComputedSource
    {
        [DataMember(Name = "source")]
        public string? Source { get; set; }
    }

    #endregion
}
