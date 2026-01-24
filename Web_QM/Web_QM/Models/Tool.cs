using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class Tool
    {
        public long Id { get; set; }

        [Required]
        public string ToolCode { get; set; }

        [Required]
        public string ToolName { get; set; }

        public string? Type { get; set; }

        public string? Unit { get; set; }

        public string? Location { get; set; }

        [Required]
        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int InitialQty { get; set; } = 0;

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int TotalImported { get; set; } = 0;

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int TotalScrapped { get; set; } = 0;

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int TotalIssued { get; set; } = 0;

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int TotalReturned { get; set; } = 0;

        [Range(0, int.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int AvailableQty { get; set; } = 0;

        public string? Note { get; set; }

        public DateOnly? CreatedDate { get; set; }

        public DateOnly? UpdatedDate { get; set; }
    }
}
