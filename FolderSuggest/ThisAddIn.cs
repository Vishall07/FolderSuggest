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
        public static Ribbon1 RibbonInstance { get; set; }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            var ribbon = new Ribbon1();
            RibbonInstance = ribbon;
            return ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _predictionService = FolderPredictionService.TryLoad();

            var explorer = Application.ActiveExplorer();
            if (explorer != null)
            {
                explorer.SelectionChange += OnSelectionChange;
            }
        }

        private void OnSelectionChange()
        {
            try
            {
                if (_predictionService == null)
                    return;

                var explorer = Application.ActiveExplorer();
                if (explorer?.Selection?.Count == 0)
                    return;

                if (explorer.Selection[1] is Outlook.MailItem mail)
                {
                    var preds = _predictionService.Predict(mail.Subject ?? "", mail.SenderEmailAddress ?? "");
                    RibbonInstance?.UpdatePredictions(preds);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in OnSelectionChange: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
