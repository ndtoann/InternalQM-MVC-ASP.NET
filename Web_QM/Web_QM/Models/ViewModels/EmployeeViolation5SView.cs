namespace Web_QM.Models.ViewModels
{
    public class EmployeeViolation5SView
    {
        public long Id { get; set; }

        public string EmployeeCode { get; set; }

        public string Content5S { get; set; }

        public DateOnly DateMonth { get; set; }

        public int Qty { get; set; }

        public decimal Amount { get; set; }
    }
}
