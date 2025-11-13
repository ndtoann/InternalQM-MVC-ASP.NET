using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class MonthlyPayroll
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng chọn nhân viên")]
        public string EmployeeCode { get; set; }

        [Required(ErrorMessage = "Giá trị không được để trống")]
        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int ResponsibleDays { get; set; }

        [Required(ErrorMessage = "Giá trị không được để trống")]
        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int RegularHours { get; set; }

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int? RegularOvertimeHours { get; set; } = 0;

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int? SundayOvertimeHours { get; set; } = 0;

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int? HolidayOvertimeHours { get; set; } = 0;

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int? LateNightOvertimeHours { get; set; } = 0;

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int? Shift3Hours { get; set; } = 0;

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int? Shift3OvertimeHours { get; set; } = 0;

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int? Shift3SundayOvertimeHours { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? BusinessTripAndPhoneFee { get; set; } = 0;

        [Required(ErrorMessage = "Giá trị không được để trống")]
        public string DateMonth { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? Penalty5S { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? UnionFee { get; set; } = 0;

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? PIT { get; set; } = 0;

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int? PaidLeaveDay { get; set; }

        public decimal? AdvanceSalary { get; set; }

        public decimal? TotalSalary { get; set; }

        public int? Status { get; set; }

        public string? CreatedDate { get; set; }

        public string? UpdatedDate { get; set; }


        public int? LevelSalary { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? BaseSalary { get; set; }

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
    }
}
