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
        public Dictionary<string, CategoryConfig> Rules { get; private set; } = new();

        public bool LoadConfig()
        {
            try
            {
                var assemblyPath = Assembly.GetExecutingAssembly().Location;
                var assemblyDir = Path.GetDirectoryName(assemblyPath) ?? string.Empty;
                var configPath = Path.Combine(assemblyDir, "readiness_checklist.json");

                if (!File.Exists(configPath))
                {
                    TaskDialog.Show("FM Readiness", $"Checklist file not found at: {configPath}");
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
                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("FM Readiness", $"Failed to load checklist JSON: {ex.Message}");
                return false;
            }
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
