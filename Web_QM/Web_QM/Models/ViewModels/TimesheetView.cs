namespace Web_QM.Models.ViewModels
{
    public class TimeSheetView
    {
        public long Id { get; set; }

        public long EmployeeId { get; set; }

        public string EmployeeCode { get; set; }

        public string EmployeeName { get; set; }

        public string DateMonth { get; set; }

        public decimal TotalHours { get; set; }

        public int Status { get; set; }
    }
}
