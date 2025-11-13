using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class TestPracticeDetail
    {
        public long Id { get; set; }
        [Required]
        public long TestPracticeId { get; set; }
        [Required]
        public string OperationName { get; set; }
        [Required]
        public int PrepTimeStandard { get; set; }
        
        public int PrepTimeActual { get; set; } = 0;
        [Required]
        public int OffsetCountStandard { get; set; }

        public int OffsetCountActual { get; set; } = 0;

        public string? Note { get; set; }
    }
}
