namespace Web_QM.Models.ViewModels
{
    public class SalaryView
    {
        public long Id { get; set; }

        public decimal BaseSalary { get; set; }

        public decimal LevelSalary { get; set; }

        public decimal? InsuranceSalary { get; set; } = 0;

        public string EmployeeCode { get; set; }

        public string EmployeeName { get; set; }

        public string Department { get; set; }
    }
}
