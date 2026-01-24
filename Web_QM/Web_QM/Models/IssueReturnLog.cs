using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class IssueReturnLog
    {
        public long Id { get; set; }

        [Required]
        public long ToolId { get; set; }

        public string? Machine { get; set; }

        public string? IntendedUse { get; set; }

        public DateOnly IssuedDate { get; set; }

        [Required]
        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int IssuedQty { get; set; } = 0;

        public string IssuedStaff { get; set; }

        public string IssuedWarehouseStaff { get; set; }

        public DateOnly? ReturnDate { get; set; }

        [Range(0, int.MaxValue)]
        public int ReturnQty { get; set; } = 0;

        public string? ReturnStaff { get; set; }

        public string? ReturnWarehouseStaff { get; set; }

        public string? Note { get; set; }

        public DateOnly? CreatedDate { get; set; }

        public DateOnly? UpdatedDate { get; set; }
    }
}
