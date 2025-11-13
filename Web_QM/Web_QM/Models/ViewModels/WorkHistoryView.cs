namespace Web_QM.Models.ViewModels
{
    public class WorkHistoryView
    {
        public DateOnly StartDate { get; set; }
        public DateOnly? EndDate { get; set; }
        public string Department { get; set; }

        public int KaizenCount { get; set; }
        public int ErrorCount { get; set; }
        public int Violation5SCount { get; set; }
    }
}
