using System;

namespace FolderSuggest.Models
{
    public class EmailItem
    {
        public string Subject { get; set; }
        public string From { get; set; }
        public DateTime ReceivedTime { get; set; }
        public string FolderName { get; set; }
        public string Body { get; set; }
    }
}
