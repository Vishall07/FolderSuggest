using Microsoft.ML.Data;

namespace FolderSuggest.Models
{
    public class EmailData
    {
        public string FolderName { get; set; }
        public string Subject { get; set; }
        public string SenderEmail { get; set; }
    }

    public class EmailPrediction
    {
        [ColumnName("PredictedLabel")]
        public string PredictedFolder { get; set; }

        public float[] Score { get; set; }
    }
}
