using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class MachineMaintenance
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng chọn mã máy")]
        public string MachineCode { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập ngày tháng")]
        public DateOnly DateMonth { get; set; }

        public string? MaintenanceStaff { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập nội dung bảo dưỡng")]
        public string MaintenanceContent { get; set; }

        public int IsComplete { get; set; } = 0;

        public string? Note { get; set; }
    }
}
