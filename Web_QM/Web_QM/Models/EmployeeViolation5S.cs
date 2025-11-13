using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class EmployeeViolation5S
    {
        public long Id { get; set; }

        [Required]
        public string EmployeeCode { get; set; }

        [Required]
        public long Violation5SId { get; set; }

        [Required]
        public DateOnly DateMonth {  get; set; }

        [Required]
        public int Qty { get; set; }
    }
}
