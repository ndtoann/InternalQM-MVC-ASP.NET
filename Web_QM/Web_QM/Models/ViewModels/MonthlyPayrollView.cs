namespace Web_QM.Models.ViewModels
{
    public class MonthlyPayrollView
    {
        public long Id { get; set; }

        public string EmployeeCode { get; set; }

        public string EmployeeName { get; set; }

        public string Department { get; set; }

        public int ResponsibleDays { get; set; }

        public int RegularHours { get; set; }

        public int RegularOvertimeHours { get; set; }

        public int SundayOvertimeHours { get; set; }

        public int HolidayOvertimeHours { get; set; }

        public int LateNightOvertimeHours { get; set; }

        public decimal TotalSalary { get; set; }

        public string DateMonth { get; set; }

        public int Status { get; set; }
    }
}
