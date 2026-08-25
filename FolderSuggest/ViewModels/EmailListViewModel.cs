using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using FolderSuggest.Models;
using FolderSuggest.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FolderSuggest.ViewModels
{
    public class EmailListViewModel : INotifyPropertyChanged
    {
        private ObservableCollection<EmailItem> _emails;
        private bool _isLoading;
        private string _statusMessage;
        private readonly OutlookService _outlookService;

        public ObservableCollection<EmailItem> Emails
        {
            get => _emails;
            set
            {
                if (_emails != value)
                {
                    _emails = value;
                    OnPropertyChanged(nameof(Emails));
                }
            }
        }

        public bool IsLoading
        {
            get => _isLoading;
            set
            {
                if (_isLoading != value)
                {
                    _isLoading = value;
                    OnPropertyChanged(nameof(IsLoading));
                }
            }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                if (_statusMessage != value)
                {
                    _statusMessage = value;
                    OnPropertyChanged(nameof(StatusMessage));
                }
            }
        }

        public ICommand LoadEmailsCommand { get; }

        public EmailListViewModel(Outlook.Application outlookApp)
        {
            _emails = new ObservableCollection<EmailItem>();
            _outlookService = new OutlookService(outlookApp);
            LoadEmailsCommand = new RelayCommand(LoadEmails);
        }

        public void LoadEmails()
        {
            IsLoading = true;
            StatusMessage = "Loading emails...";

            try
            {
                var emailList = _outlookService.GetEmailsFromFolders();
                Emails.Clear();

                foreach (var email in emailList.OrderByDescending(e => e.ReceivedTime))
                {
                    Emails.Add(email);
                }

                StatusMessage = $"Loaded {Emails.Count} emails";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;

        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter) => _canExecute?.Invoke() ?? true;

        public void Execute(object parameter) => _execute?.Invoke();
    }
}
