using System;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
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

                WebView.Source = new Uri(indexPath);

                _isWebViewReady = true;

                WebViewPaneController.RegisterPane(this);

                if (!string.IsNullOrEmpty(_pendingJson))
                {
                    PostAuditResults(_pendingJson);
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

                if (message.IndexOf("get2dViews", StringComparison.OrdinalIgnoreCase) >= 0
                    && TryExtractInt(message, "elementId", out var elementId))
                {
                    WebViewPaneController.Request2dViews(elementId);
                    return;
                }

                if (message.IndexOf("open2dView", StringComparison.OrdinalIgnoreCase) >= 0
                    && TryExtractInt(message, "viewId", out var viewId))
                {
                    WebViewPaneController.RequestOpen2dView(viewId);
                    return;
                }

                if (message.IndexOf("selectZoom", StringComparison.OrdinalIgnoreCase) >= 0
                    && TryExtractInt(message, "elementId", out var selectId))
                {
                    WebViewPaneController.RequestSelectZoom(selectId);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WebMessageReceived error: {ex.Message}");
            }
        }

        private static bool TryExtractInt(string json, string key, out int value)
        {
            value = 0;
            var match = Regex.Match(json, $"\"{key}\"\\s*:\\s*(\\d+)");
            return match.Success && int.TryParse(match.Groups[1].Value, out value);
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
