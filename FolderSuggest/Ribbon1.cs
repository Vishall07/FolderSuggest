using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using FolderSuggest.Views;

namespace FolderSuggest
{
    [ComVisible(true)]
    public partial class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private List<(string FolderName, float Score)> _predictions = new List<(string, float)>();

        public string GetCustomUI(string RibbonID)
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream($"{assembly.GetName().Name}.Ribbon1.xml"))
            using (var reader = new System.IO.StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        public void UpdatePredictions(List<(string FolderName, float Score)> predictions)
        {
            _predictions = predictions ?? new List<(string, float)>();
            ribbon?.Invalidate();
        }

        public string GetFolderLabel(Office.IRibbonControl control)
        {
            try
            {
                int index = int.Parse(control.Id.Replace("Folder", "")) - 1;
                if (index < _predictions.Count)
                {
                    var (name, score) = _predictions[index];
                    return $"{name} ({score:P0})";
                }
                return "—";
            }
            catch
            {
                return "—";
            }
        }

        public bool GetFolderEnabled(Office.IRibbonControl control)
        {
            try
            {
                int index = int.Parse(control.Id.Replace("Folder", "")) - 1;
                return index < _predictions.Count;
            }
            catch
            {
                return false;
            }
        }

        public void FolderButtonClick(Office.IRibbonControl control)
        {
            if (control.Id == "Folder1" || control.Id == "Folder2" || control.Id == "Folder3")
            {
                HandlePredictionFolderClick(control);
            }
            else if (control.Id == "Folder5")
            {
                var outlookApp = Globals.ThisAddIn.Application;
                var window = new TrainingWindow(outlookApp);
                window.Show();
            }
        }

        private void HandlePredictionFolderClick(Office.IRibbonControl control)
        {
            try
            {
                int index = int.Parse(control.Id.Replace("Folder", "")) - 1;
                if (index >= _predictions.Count)
                    return;

                var targetFolderName = _predictions[index].FolderName;
                MoveSelectedEmailToFolder(targetFolderName);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error moving email: {ex.Message}");
            }
        }

        private void MoveSelectedEmailToFolder(string folderName)
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer?.Selection?.Count == 0)
                    return;

                if (explorer.Selection[1] is Outlook.MailItem mail)
                {
                    var ns = Globals.ThisAddIn.Application.GetNamespace("MAPI");
                    var inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                    var target = FindSubFolder(inbox, folderName);
                    if (target != null)
                    {
                        mail.Move(target);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in MoveSelectedEmailToFolder: {ex.Message}");
            }
        }

        private Outlook.MAPIFolder FindSubFolder(Outlook.MAPIFolder parent, string targetName)
        {
            try
            {
                foreach (Outlook.MAPIFolder folder in parent.Folders)
                {
                    if (folder.Name.Equals(targetName, StringComparison.OrdinalIgnoreCase))
                        return folder;

                    var found = FindSubFolder(folder, targetName);
                    if (found != null)
                        return found;
                }
            }
            catch
            {
            }

            return null;
        }
    }
}
