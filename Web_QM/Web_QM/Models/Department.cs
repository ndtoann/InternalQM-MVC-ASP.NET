using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class Department
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập tên bộ phận")]
        public string DepartmentName { get; set; }

        public string? Note { get; set; }
    }
}
