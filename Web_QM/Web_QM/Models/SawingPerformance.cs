using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class SawingPerformance
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập mã nhân viên")]
        public string EmployeeCode { get; set; }

        public string? EmployeeName { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập doanh số")]
        [Range(0, 999999999, ErrorMessage = "Doanh số không hợp lệ")]
        public decimal SalesAmountUSD { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập thời gian làm việc")]
        [Range(0, 999999999, ErrorMessage = "Thời gian không hợp lệ")]
        public int WorkMinute { get; set; }

        public decimal? SalesRate { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập số năm")]
        [Range(2010, 2100, ErrorMessage = "Số năm không hợp lệ")]
        public int MeasurementYear { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập số tháng")]
        [Range(1, 12, ErrorMessage = "Số tháng không hợp lệ")]
        public int MeasurementMonth { get; set; }
    }
}
