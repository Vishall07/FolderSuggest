using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using FolderSuggest.Services;

namespace FolderSuggest
{
    public partial class ThisAddIn
    {
        private FolderPredictionService _predictionService;
        private Dictionary<Outlook.Explorer, Outlook.ExplorerEvents_10_SelectionChangeEventHandler> _explorerHandlers;
        public static Ribbon1 RibbonInstance { get; set; }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            var ribbon = new Ribbon1();
            RibbonInstance = ribbon;
            return ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("=== FolderSuggest AddIn Starting ===");

            _explorerHandlers = new Dictionary<Outlook.Explorer, Outlook.ExplorerEvents_10_SelectionChangeEventHandler>();
            _predictionService = FolderPredictionService.TryLoad();

            System.Diagnostics.Debug.WriteLine($"Prediction service loaded: {(_predictionService != null)}");

            try
            {
                var explorers = Application.Explorers;
                System.Diagnostics.Debug.WriteLine($"Number of explorers: {explorers.Count}");

                foreach (Outlook.Explorer explorer in explorers)
                {
                    HookExplorer(explorer);
                }

                Application.Explorers.NewExplorer += Application_NewExplorer;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error during startup: {ex.Message}\n{ex.StackTrace}");
            }
        }

        private void Application_NewExplorer(Outlook.Explorer explorer)
        {
            System.Diagnostics.Debug.WriteLine("New explorer opened, hooking event...");
            HookExplorer(explorer);
        }

        private void HookExplorer(Outlook.Explorer explorer)
        {
            try
            {
                if (explorer == null || _explorerHandlers.ContainsKey(explorer))
                    return;

                var handler = new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(OnSelectionChange);
                explorer.SelectionChange += handler;
                _explorerHandlers[explorer] = handler;

                System.Diagnostics.Debug.WriteLine($"Successfully hooked explorer. Total hooked: {_explorerHandlers.Count}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error hooking explorer: {ex.Message}");
            }
        }

        private void OnSelectionChange()
        {
            System.Diagnostics.Debug.WriteLine("OnSelectionChange called");

            try
            {
                if (_predictionService == null)
                {
                    System.Diagnostics.Debug.WriteLine("Prediction service is null, skipping");
                    return;
                }

                var explorer = Application.ActiveExplorer();
                if (explorer == null)
                {
                    System.Diagnostics.Debug.WriteLine("No active explorer");
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"Selection count: {explorer.Selection.Count}");

                if (explorer.Selection.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("No selection");
                    return;
                }

                var selectedItem = explorer.Selection[1];
                System.Diagnostics.Debug.WriteLine($"Selected item type: {selectedItem.GetType().Name}");

                if (selectedItem is Outlook.MailItem mail)
                {
                    System.Diagnostics.Debug.WriteLine($"Processing email: {mail.Subject}");
                    var preds = _predictionService.Predict(mail.Subject ?? "", mail.SenderEmailAddress ?? "");
                    System.Diagnostics.Debug.WriteLine($"Predictions received: {preds.Count}");

                    RibbonInstance?.UpdatePredictions(preds);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Selected item is not a MailItem: {selectedItem.GetType().Name}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in OnSelectionChange: {ex.Message}\n{ex.StackTrace}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                if (_explorerHandlers != null)
                {
                    foreach (var kvp in _explorerHandlers)
                    {
                        kvp.Key.SelectionChange -= kvp.Value;
                    }
                    _explorerHandlers.Clear();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error during shutdown: {ex.Message}");
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
