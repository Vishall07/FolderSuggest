using System;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Threading;
using FolderSuggest.Services;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FolderSuggest.ViewModels
{
    public class TrainingViewModel : INotifyPropertyChanged
    {
        private int _progressValue;
        private string _statusMessage;
        private bool _isTraining;
        private readonly Outlook.Application _outlookApp;
        private readonly Dispatcher _dispatcher;

        public int ProgressValue
        {
            get => _progressValue;
            set
            {
                if (_progressValue != value)
                {
                    _progressValue = value;
                    OnPropertyChanged(nameof(ProgressValue));
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

        public bool IsTraining
        {
            get => _isTraining;
            set
            {
                if (_isTraining != value)
                {
                    _isTraining = value;
                    OnPropertyChanged(nameof(IsTraining));
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        public ICommand StartTrainingCommand { get; }

        public TrainingViewModel(Outlook.Application outlookApp, Dispatcher dispatcher)
        {
            _outlookApp = outlookApp;
            _dispatcher = dispatcher ?? Dispatcher.CurrentDispatcher;
            StatusMessage = "Ready. Click 'Start Training' to begin.";
            StartTrainingCommand = new RelayCommand(StartTraining, () => !IsTraining);
        }

        private async void StartTraining()
        {
            IsTraining = true;
            ProgressValue = 0;

            var trainer = new FolderPredictionTrainer(_outlookApp);
            trainer.ProgressUpdated += OnProgressUpdated;

            try
            {
                await Task.Run(() => trainer.TrainModel());
            }
            catch (Exception ex)
            {
                StatusMessage = $"Training failed: {ex.Message}";
                ProgressValue = 0;
            }
            finally
            {
                trainer.ProgressUpdated -= OnProgressUpdated;
                IsTraining = false;
            }
        }

        private void OnProgressUpdated(int progress, string message)
        {
            _dispatcher.Invoke(() =>
            {
                ProgressValue = progress;
                StatusMessage = message;
            });
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
