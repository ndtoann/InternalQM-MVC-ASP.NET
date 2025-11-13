using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class Machine_MG
    {
        [Required]
        public string MachineCode { get; set; }

        [Required]
        public long MachineGroupId { get; set; }

        [Required]
        public string Material { get; set; }
    }
}
