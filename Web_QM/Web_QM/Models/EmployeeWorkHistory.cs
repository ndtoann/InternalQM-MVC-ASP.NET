using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class EmployeeWorkHistory
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập mã nhân viên")]
        public string EmployeeCode { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập tên nhân viên")]
        public string EmployeeName { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập bộ phận")]
        public string Department {  get; set; }

        [Required(ErrorMessage = "Vui lòng nhập ngày bắt đầu")]
        public DateOnly StartDate { get; set; }

        public DateOnly? EndDate { get; set; }

        public string? Note { get; set; }
    }
}
