using System;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.UI;
using FMReadiness_v3.UI.Panes;
using Microsoft.Web.WebView2.Core;

namespace FMReadiness_v3.UI
{
    public partial class AuditWebPane : UserControl
    {
        private bool _isWebViewReady;
        private string? _pendingJson;

        public AuditWebPane()
        {
            InitializeComponent();
        }

        private async void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // Use a writable user data folder because add-ins often run from restricted paths.
                var userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "FMReadiness_v3",
                    "WebView2");

                Directory.CreateDirectory(userDataFolder);

                var env = await CoreWebView2Environment.CreateAsync(
                    browserExecutableFolder: null,
                    userDataFolder: userDataFolder,
                    options: null);

                await WebView.EnsureCoreWebView2Async(env);

                WebView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                WebView.CoreWebView2.Settings.AreDevToolsEnabled = true;
                WebView.CoreWebView2.Settings.IsStatusBarEnabled = false;

                WebView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;

                var assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;
                var indexPath = Path.Combine(assemblyDir, "UI", "index.html");
                if (!File.Exists(indexPath))
                    indexPath = Path.Combine(assemblyDir, "ui", "index.html");

                if (!File.Exists(indexPath))
                {
                    TaskDialog.Show("FM Readiness", $"HTML file not found.\nLooked in: {assemblyDir}\\UI\\");
                    return;
                }

                var indexUri = new Uri(indexPath);
                var cacheBuster = File.GetLastWriteTimeUtc(indexPath).Ticks;
                var uriBuilder = new UriBuilder(indexUri)
                {
                    Query = $"v={cacheBuster}"
                };
                WebView.Source = uriBuilder.Uri;

                _isWebViewReady = true;

                WebViewPaneController.RegisterPane(this);

                var pendingJson = _pendingJson;
                if (!string.IsNullOrEmpty(pendingJson))
                {
                    PostAuditResults(pendingJson!);
                    _pendingJson = null;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show(
                    "FM Readiness",
                    "WebView2 initialization failed:\n" + ex.Message +
                    "\n\nTip: Ensure Microsoft Edge WebView2 Runtime is installed and that %LOCALAPPDATA% is writable.");
            }
        }

        private void CoreWebView2_WebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                var message = e.WebMessageAsJson;
                if (string.IsNullOrWhiteSpace(message)) return;

                using var doc = JsonDocument.Parse(message);
                var root = doc.RootElement;

                if (root.TryGetProperty("action", out var actionProp))
                {
                    var action = actionProp.GetString();

                    if (string.Equals(action, "get2dViews", StringComparison.OrdinalIgnoreCase)
                        && root.TryGetProperty("elementId", out var elementIdProp)
                        && elementIdProp.TryGetInt32(out var elementId))
                    {
                        WebViewPaneController.Request2dViews(elementId);
                        return;
                    }

                    if (string.Equals(action, "open2dView", StringComparison.OrdinalIgnoreCase)
                        && root.TryGetProperty("viewId", out var viewIdProp)
                        && viewIdProp.TryGetInt32(out var viewId))
                    {
                        WebViewPaneController.RequestOpen2dView(viewId);
                        return;
                    }

                    if (string.Equals(action, "getSelectedElements", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.GetSelectedElements);
                        return;
                    }

                    if (string.Equals(action, "getCategoryStats", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.GetCategoryStats,
                            message);
                        return;
                    }

                    if (string.Equals(action, "setInstanceParams", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.SetInstanceParams,
                            message);
                        return;
                    }

                    if (string.Equals(action, "setCategoryParams", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.SetCategoryParams,
                            message);
                        return;
                    }

                    if (string.Equals(action, "setTypeParams", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.SetTypeParams,
                            message);
                        return;
                    }

                    if (string.Equals(action, "copyComputedToParam", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.CopyComputedToParam,
                            message);
                        return;
                    }

                    if (string.Equals(action, "refreshAudit", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestAuditRefresh();
                        return;
                    }

                    // COBie Preset Operations
                    if (string.Equals(action, "getAvailablePresets", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.GetAvailablePresets);
                        return;
                    }

                    if (string.Equals(action, "loadPreset", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.LoadPreset,
                            message);
                        return;
                    }

                    if (string.Equals(action, "getPresetFields", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.GetPresetFields,
                            message);
                        return;
                    }

                    if (string.Equals(action, "setCobieFieldValues", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.SetCobieFieldValues,
                            message);
                        return;
                    }

                    if (string.Equals(action, "validateCobieReadiness", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.ValidateCobieReadiness,
                            message);
                        return;
                    }

                    if (string.Equals(action, "ensureCobieParameters", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.EnsureCobieParameters,
                            message);
                        return;
                    }

                    // COBie Editor Operations
                    if (string.Equals(action, "setCobieTypeFieldValues", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.SetCobieTypeFieldValues,
                            message);
                        return;
                    }

                    if (string.Equals(action, "setCobieCategoryFieldValues", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.SetCobieCategoryFieldValues,
                            message);
                        return;
                    }

                    if (string.Equals(action, "copyComputedToCobieParam", StringComparison.OrdinalIgnoreCase))
                    {
                        WebViewPaneController.RequestParameterEditorOperation(
                            ExternalEvents.ParameterEditorExternalEventHandler.OperationType.CopyComputedToCobieParam,
                            message);
                        return;
                    }
                }

                if (root.TryGetProperty("type", out var typeProp)
                    && string.Equals(typeProp.GetString(), "selectZoom", StringComparison.OrdinalIgnoreCase)
                    && root.TryGetProperty("elementId", out var selectIdProp)
                    && selectIdProp.TryGetInt32(out var selectId))
                {
                    WebViewPaneController.RequestSelectZoom(selectId);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WebMessageReceived error: {ex.Message}");
            }
        }

        public void PostAuditResults(string json)
        {
            if (!_isWebViewReady || WebView.CoreWebView2 == null)
            {
                _pendingJson = json;
                return;
            }

            try
            {
                WebView.CoreWebView2.PostWebMessageAsJson(json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PostAuditResults error: {ex.Message}");
            }
        }
    }
}
