using Microsoft.ML.Data;

namespace FolderSuggest.Models
{
    public class EmailData
    {
        [LoadColumn(0)]
        public string FolderName { get; set; }

        [LoadColumn(1)]
        public string Subject { get; set; }

        [LoadColumn(2)]
        public string BodyPreview { get; set; }

        [LoadColumn(3)]
        public string From { get; set; }
    }

    public class EmailPrediction
    {
        [ColumnName("PredictedLabel")]
        public string PredictedFolder { get; set; }

        public float[] Score { get; set; }
    }
}
