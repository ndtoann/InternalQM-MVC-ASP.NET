using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class MachineParameter
    {
        public long Id { get; set; }

        [Required]
        public string MachineCode { get; set; }

        [Required]
        public string Type { get; set; }

        [Required]
        public decimal Parameters { get; set; }

        [Required]
        public DateOnly DateMonth { get; set; }
    }
}
