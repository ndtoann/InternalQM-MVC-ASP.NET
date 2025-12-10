using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class Timesheet
    {
        public long Id { get; set; }

        [Required]
        public long EmployeeId { get; set; }

        [Required]
        public string DateMonth { get; set; }

        [Required]
        public int Status { get; set; } = 1;
    }
}
