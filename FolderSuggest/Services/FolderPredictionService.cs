using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.ML;
using Microsoft.ML.Data;
using FolderSuggest.Models;

namespace FolderSuggest.Services
{
    public class FolderPredictionService
    {
        private readonly PredictionEngine<EmailData, EmailPrediction> _engine;
        private readonly string[] _labelNames;

        private static string GetModelPath()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, "FolderSuggest", "EmailFolderModel.zip");
        }

        public FolderPredictionService(PredictionEngine<EmailData, EmailPrediction> engine, string[] labelNames)
        {
            _engine = engine;
            _labelNames = labelNames;
        }

        public static FolderPredictionService TryLoad()
        {
            try
            {
                var modelPath = GetModelPath();
                if (!File.Exists(modelPath))
                    return null;

                var mlContext = new MLContext();
                ITransformer model = mlContext.Model.Load(modelPath, out var schema);
                var engine = mlContext.Model.CreatePredictionEngine<EmailData, EmailPrediction>(model);

                var scoreCol = schema["Score"];
                VBuffer<ReadOnlyMemory<char>> slotNames = default;
                scoreCol.Annotations.GetValue("SlotNames", ref slotNames);
                var labelNames = slotNames.DenseValues().Select(v => v.ToString()).ToArray();

                return new FolderPredictionService(engine, labelNames);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading prediction model: {ex.Message}");
                return null;
            }
        }

        public List<(string FolderName, float Score)> Predict(string subject, string senderEmail, int topN = 3)
        {
            try
            {
                var input = new EmailData
                {
                    Subject = subject ?? "",
                    SenderEmail = senderEmail ?? ""
                };

                var prediction = _engine.Predict(input);

                var results = _labelNames
                    .Select((name, index) => (name, score: prediction.Score[index]))
                    .OrderByDescending(x => x.score)
                    .Take(topN)
                    .ToList();

                return results;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error predicting folder: {ex.Message}");
                return new List<(string, float)>();
            }
        }
    }
}
