using System.Windows;
using FolderSuggest.ViewModels;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FolderSuggest.Views
{
    public partial class EmailListWindow : Window
    {
        public EmailListWindow(Outlook.Application outlookApp)
        {
            InitializeComponent();
            var viewModel = new EmailListViewModel(outlookApp);
            DataContext = viewModel;
            viewModel.LoadEmails();
        }
    }
}
