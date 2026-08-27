using System.Windows;
using FolderSuggest.ViewModels;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FolderSuggest.Views
{
    public partial class TrainingWindow : Window
    {
        public TrainingWindow(Outlook.Application outlookApp)
        {
            InitializeComponent();
            DataContext = new TrainingViewModel(outlookApp);
        }
    }
}
