using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class ProcessStep
    {
        public long Id { get; set; }

        [Required]
        public long ProductionProcessId { get; set; }

        [Required]
        [Range(1, int.MaxValue)]
        public int StepNumber { get; set; }

        [Required]
        public string Department { get; set; }

        public string? Content { get; set; }

        public string? Fixture { get; set; }

        [Required]
        [Range(0, (double)decimal.MaxValue)]
        public decimal EstimatedTime { get; set; }

        public string? QtyPerSet { get; set; } = "1";

        public string? Picture { get; set; }

        public string? Note { get; set; }
    }
}
