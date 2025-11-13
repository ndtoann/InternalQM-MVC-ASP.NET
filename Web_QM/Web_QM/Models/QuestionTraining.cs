using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class QuestionTraining
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
        [Required]
        public long ExamTrainingId { get; set; }

        public int? DisplayOrder { get; set; }

        public string? CreatedDate { get; set; }

        public string? UpdatedDate { get; set; }
    }
}
