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

        private const string ModelPath = @"EmailFolderModel.zip";

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
                mlContext.Model.Save(model, dataView.Schema, ModelPath);

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
                            var emailData = new EmailData
                            {
                                FolderName = folder.Name,
                                Subject = mailItem.Subject ?? "",
                                BodyPreview = (mailItem.Body ?? "").Substring(0, Math.Min(500, (mailItem.Body ?? "").Length)),
                                From = mailItem.SenderName ?? ""
                            };
                            trainingData.Add(emailData);
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
            var pipeline = mlContext.Transforms.Text.FeaturizeText("SubjectFeatures", new TextFeaturizingEstimator.Options
            {
                OutputColumnName = "SubjectFeatures",
                InputColumnName = "Subject"
            })
            .Append(mlContext.Transforms.Text.FeaturizeText("BodyFeatures", new TextFeaturizingEstimator.Options
            {
                OutputColumnName = "BodyFeatures",
                InputColumnName = "BodyPreview"
            }))
            .Append(mlContext.Transforms.Concatenate("Features", "SubjectFeatures", "BodyFeatures"))
            .Append(mlContext.MulticlassClassification.Trainers.FastTree(
                labelColumnName: "FolderName",
                featureColumnName: "Features"));

            return pipeline;
        }

        public static bool ModelExists()
        {
            return File.Exists(ModelPath);
        }

        public static string GetModelPath()
        {
            return Path.GetFullPath(ModelPath);
        }
    }
}
