using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class ExamTraining
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập tên")]
        public string ExamName { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập thời gian làm bài")]
        [Range(1, 1000, ErrorMessage = "Số phút lớn hơn 1 và nhỏ hơn 1000")]
        public int DurationMinute { get; set; }

        public string? EssayQuestion { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập tổng điểm bài thi tự luận")]
        public int TlTotal { get; set; }

        public int IsActive { get; set; } = 0;

        public DateOnly? CreatedDate { get; set; }

        public DateOnly? UpdatedDate { get; set; }
    }
}
