using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class ToolSupplyLog
    {
        public long Id { get; set; }

        [Required]
        public long ToolId { get; set; }

        [Required]
        public string Type { get; set; }

        [Required]
        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int Qty { get; set; } = 0;

        public string? IntendedUse { get; set; }

        public string? Describe { get; set; }

        public string? WarehouseStaff { get; set; }

        public string? HandOverStaff { get; set; }

        [Required]
        public DateOnly DateMonth { get; set; }

        public string? Note { get; set; }

        public DateOnly? CreatedDate { get; set; }

        public DateOnly? UpdatedDate { get; set; }
    }
}
