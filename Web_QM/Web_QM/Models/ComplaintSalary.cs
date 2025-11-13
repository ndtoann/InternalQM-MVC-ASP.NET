using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class ComplaintSalary
    {
        public long Id { get; set; }

        [Required]
        public long MonthlyPayId { get; set; }

        [Required]
        public string CompaintContent { get; set; } = null!;

        public string? CreatedDate { get; set; }

        public string? UpdatedDate { get; set; }
    }
}
