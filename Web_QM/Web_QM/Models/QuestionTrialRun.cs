using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class QuestionTrialRun
    {
        public long Id { get; set; }
        [Required(ErrorMessage = "Vui lòng nhập câu hỏi!")]
        public string QuestionText { get; set; }
        [Required(ErrorMessage = "Vui lòng nhập đáp án")]
        public string OptionA { get; set; }
        [Required(ErrorMessage = "Vui lòng nhập đáp án")]
        public string OptionB { get; set; }
        [Required(ErrorMessage = "Vui lòng nhập đáp án")]
        public string OptionC { get; set; }
        [Required(ErrorMessage = "Vui lòng nhập đáp án")]
        public string OptionD { get; set; }
        [Required(ErrorMessage = "Vui lòng chọn đáp án đúng!")]
        public string CorrectOption { get; set; }

        public int IsCritical { get; set; } = 0;
        [Required]
        public long ExamTrialRunId { get; set; }

        public int? DisplayOrder { get; set; }

        public string? CreatedDate { get; set; }

        public string? UpdatedDate { get; set; }
    }
}
