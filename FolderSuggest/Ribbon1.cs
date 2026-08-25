using System;
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

        public void FolderButtonClick(Office.IRibbonControl control)
        {
            // Placeholder: Folder button clicked
            // TODO: Wire backend to navigate to selected folder
        }

    }
}
