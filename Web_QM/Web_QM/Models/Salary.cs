using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class Salary
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng chọn nhân viên")]
        public string EmployeeCode { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập bậc lương")]
        public int LevelSalary { get; set; }

        [Required(ErrorMessage = "Giá trị không được để trống")]
        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal BaseSalary { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? InsuranceSalary { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? MealAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? DailyResponsibilityPay { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? FuelAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? ExternalWorkAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? HousingSubsidy { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? DiligencePay { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? NoViolationBonus { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? HazardousAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? CNCStressAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? SeniorityAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? CertificateAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? WorkingEnvironmentAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? JobPositionAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? MachineTestRunAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? RVFMachineMeasurementAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? StainlessSteelCleaningAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? FactoryGuardAllowance { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? HeavyDutyAllowance { get; set; } = 0;

        public string? CreatedDate { get; set; }

        public string? UpdatedDate { get; set; }
    }
}
