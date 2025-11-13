using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class ReplacementEquipmentAndSupplies
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập thiết bị, vật tư thay thế")]
        public string EquipmentAndlSupplies { get; set; }

        public string? FilePdf { get; set; }

        public long EquipmentRepairId { get; set; }
    }
}
