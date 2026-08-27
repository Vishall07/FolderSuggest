using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.ML;
using Microsoft.ML.Data;
using FolderSuggest.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FolderSuggest.Services
{
    public class FolderPredictionTrainer
    {
        private readonly Outlook.Application _outlookApp;
        public event Action<int, string> ProgressUpdated;

        private static string GetModelPath()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, "FolderSuggest", "EmailFolderModel.zip");
        }

        public FolderPredictionTrainer(Outlook.Application outlookApp)
        {
            _outlookApp = outlookApp;
        }

        public void TrainModel()
        {
            try
            {
                ProgressUpdated?.Invoke(10, "Collecting email data from folders...");
                var trainingData = CollectTrainingData();

                if (trainingData.Count == 0)
                {
                    ProgressUpdated?.Invoke(0, "No emails found to train on.");
                    return;
                }

                ProgressUpdated?.Invoke(30, $"Collected {trainingData.Count} emails. Building ML pipeline...");
                var mlContext = new MLContext();

                var dataView = mlContext.Data.LoadFromEnumerable(trainingData);
                var pipeline = BuildPipeline(mlContext);

                ProgressUpdated?.Invoke(50, "Training model...");
                var model = pipeline.Fit(dataView);

                ProgressUpdated?.Invoke(80, "Saving model...");
                var modelPath = GetModelPath();
                Directory.CreateDirectory(Path.GetDirectoryName(modelPath));
                mlContext.Model.Save(model, dataView.Schema, modelPath);

                ProgressUpdated?.Invoke(100, "Training complete! Model saved.");
            }
            catch (Exception ex)
            {
                ProgressUpdated?.Invoke(0, $"Error: {ex.Message}");
                throw;
            }
        }

        private List<EmailData> CollectTrainingData()
        {
            var trainingData = new List<EmailData>();

            try
            {
                var ns = _outlookApp.GetNamespace("MAPI");
                var rootFolder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                CollectEmailsRecursive(rootFolder, trainingData, isRootInbox: true);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error collecting training data: {ex.Message}");
            }

            return trainingData;
        }

        private void CollectEmailsRecursive(Outlook.MAPIFolder folder, List<EmailData> trainingData, bool isRootInbox = false)
        {
            try
            {
                if (isRootInbox || folder.Name.Equals("Inbox", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (Outlook.MAPIFolder subFolder in folder.Folders)
                    {
                        CollectEmailsRecursive(subFolder, trainingData, isRootInbox: false);
                    }
                }
                else
                {
                    var items = folder.Items;
                    foreach (object item in items)
                    {
                        if (item is Outlook.MailItem mailItem)
                        {
                            trainingData.Add(new EmailData
                            {
                                FolderName = folder.Name,
                                Subject = mailItem.Subject ?? "",
                                SenderEmail = mailItem.SenderEmailAddress ?? ""
                            });
                        }
                    }

                    foreach (Outlook.MAPIFolder subFolder in folder.Folders)
                    {
                        CollectEmailsRecursive(subFolder, trainingData, isRootInbox: false);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing folder {folder.Name}: {ex.Message}");
            }
        }

        private IEstimator<ITransformer> BuildPipeline(MLContext mlContext)
        {
            return mlContext.Transforms.Conversion.MapValueToKey("Label", "FolderName")
                .Append(mlContext.Transforms.Text.FeaturizeText("SubjectFeatures", "Subject"))
                .Append(mlContext.Transforms.Text.FeaturizeText("SenderFeatures", "SenderEmail"))
                .Append(mlContext.Transforms.Concatenate("Features", "SubjectFeatures", "SenderFeatures"))
                .Append(mlContext.MulticlassClassification.Trainers.SdcaMaximumEntropy(
                    labelColumnName: "Label",
                    featureColumnName: "Features"))
                .Append(mlContext.Transforms.Conversion.MapKeyToValue("PredictedLabel", "PredictedLabel"));
        }

        public static bool ModelExists()
        {
            return File.Exists(GetModelPath());
        }
    }
}
