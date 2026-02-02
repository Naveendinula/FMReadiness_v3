using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using RevitApp = Autodesk.Revit.ApplicationServices.Application;

namespace FMReadiness_v3.Services
{
    public class CobieParameterService
    {
        public class EnsureParametersResult
        {
            public int Created { get; set; }
            public int UpdatedBindings { get; set; }
            public int Skipped { get; set; }
            public int Removed { get; set; }
            public List<string> Warnings { get; } = new();
        }

        public EnsureParametersResult EnsureParameters(
            UIApplication uiApp,
            Document doc,
            CobiePreset preset,
            bool includeAliases,
            bool removeCobieParameters)
        {
            var result = new EnsureParametersResult();

            if (doc.IsFamilyDocument)
            {
                result.Warnings.Add("Family documents are not supported for parameter binding.");
                return result;
            }

            if (doc.IsReadOnly)
            {
                result.Warnings.Add("Document is read-only. Parameters were not added.");
                return result;
            }

            if (preset.Tables == null)
            {
                result.Warnings.Add("No preset tables found.");
                return result;
            }

            RevitApp app = uiApp.Application;
            var originalSharedFile = app.SharedParametersFilename;

            var sharedParamPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "FMReadiness_v3",
                "SharedParameters",
                "FMReadiness_SharedParams.txt");

            Directory.CreateDirectory(Path.GetDirectoryName(sharedParamPath) ?? string.Empty);
            if (!File.Exists(sharedParamPath))
                File.WriteAllText(sharedParamPath, string.Empty);

            app.SharedParametersFilename = sharedParamPath;
            var defFile = app.OpenSharedParameterFile();
            if (defFile == null)
            {
                result.Warnings.Add("Could not open shared parameter file.");
                app.SharedParametersFilename = originalSharedFile;
                return result;
            }

            var group = defFile.Groups.get_Item("COBie") ?? defFile.Groups.Create("COBie");

            try
            {
                using var tx = new Transaction(doc, "Ensure COBie Parameters");
                tx.Start();

                if (removeCobieParameters)
                {
                    result.Removed = RemoveParametersByPrefix(doc, "COBie.");
                }

                foreach (var tableEntry in preset.Tables)
                {
                    var tableName = tableEntry.Key;
                    var table = tableEntry.Value;
                    if (table?.Fields == null) continue;

                    var categories = GetCategoriesForTable(tableName, preset);
                    if (categories.Count == 0) continue;

                    var categorySet = BuildCategorySet(app, doc, categories);
                    if (categorySet == null) continue;

                    foreach (var field in table.Fields)
                    {
                        if (field == null) continue;

                        var paramNames = new List<string>();
                        if (!string.IsNullOrWhiteSpace(field.RevitParam))
                            paramNames.Add(field.RevitParam);

                        if (includeAliases && field.AliasParams != null)
                        {
                            foreach (var alias in field.AliasParams)
                            {
                                if (string.IsNullOrWhiteSpace(alias)) continue;
                                if (alias.StartsWith("FM_", StringComparison.OrdinalIgnoreCase)
                                    || alias.StartsWith("COBie.", StringComparison.OrdinalIgnoreCase))
                                {
                                    paramNames.Add(alias);
                                }
                            }
                        }

                        if (paramNames.Count == 0) continue;
                        if (field.Computed != null && !string.IsNullOrEmpty(field.Computed.Source))
                            continue;
                        if (!string.IsNullOrEmpty(field.RevitBuiltIn))
                            continue;

                        foreach (var paramName in paramNames.Distinct(StringComparer.OrdinalIgnoreCase))
                        {
                            var bindingKind = field.Scope == "type"
                                ? ParamBindingKind.Type
                                : ParamBindingKind.Instance;

                            EnsureParameterBinding(
                                doc,
                                app,
                                group,
                                paramName,
                                field.DataType,
                                bindingKind,
                                categorySet,
                                result);
                        }
                    }
                }

                // FM Ops extensions (treat as Component/instance)
                if (preset.FmOpsExtensions?.Fields != null)
                {
                    var categories = GetCategoriesForTable("Component", preset);
                    var categorySet = BuildCategorySet(app, doc, categories);
                    if (categorySet != null)
                    {
                        foreach (var field in preset.FmOpsExtensions.Fields)
                        {
                            if (field == null) continue;
                            if (string.IsNullOrWhiteSpace(field.RevitParam)) continue;

                            var bindingKind = field.Scope == "type"
                                ? ParamBindingKind.Type
                                : ParamBindingKind.Instance;

                            EnsureParameterBinding(
                                doc,
                                app,
                                group,
                                field.RevitParam,
                                field.DataType,
                                bindingKind,
                                categorySet,
                                result);
                        }
                    }
                }

                tx.Commit();
            }
            finally
            {
                app.SharedParametersFilename = originalSharedFile;
            }

            return result;
        }

        private enum ParamBindingKind
        {
            Instance,
            Type
        }

        private void EnsureParameterBinding(
            Document doc,
            RevitApp app,
            DefinitionGroup group,
            string paramName,
            string? dataType,
            ParamBindingKind bindingKind,
            CategorySet categorySet,
            EnsureParametersResult result)
        {
            var map = doc.ParameterBindings;
            var definition = FindDefinition(map, paramName, out var existingBinding);

            if (definition == null)
            {
                definition = CreateDefinition(group, paramName, dataType);
                if (definition == null)
                {
                    result.Warnings.Add($"Failed to create definition for {paramName}");
                    return;
                }

                var newBinding = bindingKind == ParamBindingKind.Type
                    ? (Binding)app.Create.NewTypeBinding(categorySet)
                    : app.Create.NewInstanceBinding(categorySet);

#if REVIT2024_OR_GREATER
                if (map.Insert(definition, newBinding, GroupTypeId.IdentityData))
                    result.Created++;
#else
                if (map.Insert(definition, newBinding, BuiltInParameterGroup.PG_IDENTITY_DATA))
                    result.Created++;
#endif
                return;
            }

            // Existing binding: ensure type/instance matches and categories are included
            if (existingBinding != null)
            {
                var isTypeBinding = existingBinding is TypeBinding;
                if (bindingKind == ParamBindingKind.Type && !isTypeBinding)
                {
                    result.Warnings.Add($"Parameter '{paramName}' exists as instance; expected type.");
                    result.Skipped++;
                    return;
                }
                if (bindingKind == ParamBindingKind.Instance && isTypeBinding)
                {
                    result.Warnings.Add($"Parameter '{paramName}' exists as type; expected instance.");
                    result.Skipped++;
                    return;
                }

                var merged = MergeCategories(app, existingBinding.Categories, categorySet);
                if (merged != null)
                {
                    var newBinding = bindingKind == ParamBindingKind.Type
                        ? (Binding)app.Create.NewTypeBinding(merged)
                        : app.Create.NewInstanceBinding(merged);

#if REVIT2024_OR_GREATER
                    if (map.ReInsert(definition, newBinding, GroupTypeId.IdentityData))
                        result.UpdatedBindings++;
#else
                    if (map.ReInsert(definition, newBinding, BuiltInParameterGroup.PG_IDENTITY_DATA))
                        result.UpdatedBindings++;
#endif
                }
                else
                {
                    result.Skipped++;
                }
            }
            else
            {
                result.Skipped++;
            }
        }

        private Definition? CreateDefinition(DefinitionGroup group, string name, string? dataType)
        {
            try
            {
#if REVIT2024_OR_GREATER
                var specType = GetSpecTypeId(dataType);
                var options = new ExternalDefinitionCreationOptions(name, specType);
#else
                var paramType = GetParameterType(dataType);
                var options = new ExternalDefinitionCreationOptions(name, paramType);
#endif
                options.UserModifiable = true;
                options.Visible = true;
                return group.Definitions.Create(options);
            }
            catch
            {
                return null;
            }
        }

#if REVIT2024_OR_GREATER
        private ForgeTypeId GetSpecTypeId(string? dataType)
        {
            return dataType?.ToLowerInvariant() switch
            {
                "number" => SpecTypeId.Number,
                _ => SpecTypeId.String.Text
            };
        }
#else
        private ParameterType GetParameterType(string? dataType)
        {
            return dataType?.ToLowerInvariant() switch
            {
                "number" => ParameterType.Number,
                _ => ParameterType.Text
            };
        }
#endif

        private Definition? FindDefinition(BindingMap map, string name, out ElementBinding? binding)
        {
            binding = null;
            var it = map.ForwardIterator();
            it.Reset();
            while (it.MoveNext())
            {
                if (it.Key is not Definition def) continue;
                if (!def.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) continue;
                binding = it.Current as ElementBinding;
                return def;
            }
            return null;
        }

        private CategorySet? BuildCategorySet(RevitApp app, Document doc, List<BuiltInCategory> categories)
        {
            var set = app.Create.NewCategorySet();
            var added = false;
            foreach (var bic in categories)
            {
                try
                {
                    var cat = doc.Settings.Categories.get_Item(bic);
                    if (cat != null)
                    {
                        set.Insert(cat);
                        added = true;
                    }
                }
                catch
                {
                    // ignore missing categories
                }
            }
            return added ? set : null;
        }

        private CategorySet? MergeCategories(RevitApp app, CategorySet existing, CategorySet target)
        {
            var merged = app.Create.NewCategorySet();
            bool changed = false;

            foreach (Category cat in existing)
                merged.Insert(cat);

            foreach (Category cat in target)
            {
                if (!merged.Contains(cat))
                {
                    merged.Insert(cat);
                    changed = true;
                }
            }

            return changed ? merged : null;
        }

        private List<BuiltInCategory> GetCategoriesForTable(string tableName, CobiePreset preset)
        {
            if (string.Equals(tableName, "Space", StringComparison.OrdinalIgnoreCase))
            {
                return new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_Rooms,
                    BuiltInCategory.OST_MEPSpaces
                };
            }

            if (string.Equals(tableName, "Floor", StringComparison.OrdinalIgnoreCase))
            {
                return new List<BuiltInCategory> { BuiltInCategory.OST_Levels };
            }

            if (string.Equals(tableName, "Facility", StringComparison.OrdinalIgnoreCase))
            {
                return new List<BuiltInCategory> { BuiltInCategory.OST_ProjectInformation };
            }

            var list = new List<BuiltInCategory>();
            if (preset.Categories != null)
            {
                foreach (var cat in preset.Categories)
                {
                    if (Enum.TryParse(cat, out BuiltInCategory bic))
                        list.Add(bic);
                }
            }

            if (list.Count == 0)
            {
                list.Add(BuiltInCategory.OST_MechanicalEquipment);
                list.Add(BuiltInCategory.OST_DuctTerminal);
                list.Add(BuiltInCategory.OST_DuctAccessory);
                list.Add(BuiltInCategory.OST_PipeAccessory);
            }

            return list;
        }

        private int RemoveParametersByPrefix(Document doc, string prefix)
        {
            var map = doc.ParameterBindings;
            var toRemove = new List<Definition>();
            var it = map.ForwardIterator();
            it.Reset();
            while (it.MoveNext())
            {
                if (it.Key is Definition def
                    && def.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    toRemove.Add(def);
                }
            }

            int removed = 0;
            foreach (var def in toRemove)
            {
                if (map.Remove(def))
                    removed++;
            }

            return removed;
        }
    }
}
