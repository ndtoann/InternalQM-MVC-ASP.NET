using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class Violation5S
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập nội dung lỗi 5S")]
        public string Content5S { get; set; }

        public string? Note {  get; set; }

        [Required(ErrorMessage = "Vui lòng nhập số tiền phạt lỗi 5S")]
        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal Amount { get; set; }

        public string? CreatedDate { get; set; }

        public string? UpdatedDate { get; set; }
    }
}
