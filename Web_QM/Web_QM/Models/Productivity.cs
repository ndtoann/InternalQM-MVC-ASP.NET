using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class Productivity
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập mã nhân viên")]
        public string EmployeeCode { get; set; }

        public string? EmployeeName { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập năng suất")]
        [Range(0, 1000, ErrorMessage = "Năng suất không hợp lệ")]
        public decimal ProductivityScore { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập số năm")]
        [Range(2010, 2100, ErrorMessage = "Số năm không hợp lệ")]
        public int MeasurementYear { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập số tháng")]
        [Range(1, 12, ErrorMessage = "Số tháng không hợp lệ")]
        public int MeasurementMonth { get; set; }
    }
}
