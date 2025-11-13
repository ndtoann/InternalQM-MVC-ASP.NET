using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class TestPractice
    {
        public long Id { get; set; }
        [Required]
        public long EmployeeId { get; set; }
        [Required]
        public string TestName { get; set; }
        [Required]
        public string TestLevel { get; set; }
        [Required]
        public string PartName { get; set; }
        [Required]
        public string Result { get; set; }

        public string? CreatedDate { get; set; }

        public string? UpdatedDate { get; set; }

        public ICollection<TestPracticeDetail> Details { get; set; }
    }
}
