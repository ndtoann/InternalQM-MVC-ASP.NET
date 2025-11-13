using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class ExamTrialRunAnswer
    {
        public long Id { get; set; }
        [Required]
        public string EmployeeName { get; set; }
        [Required]
        public string EmployeeCode { get; set; }
        [Required]
        public string ListAnswer { get; set; }

        public int IsShow { get; set; } = 0;

        [Range(0, 100, ErrorMessage = "Điểm không hợp lệ")]
        public int MultipleChoiceCorrect { get; set; } = 0;

        [Range(0, 100, ErrorMessage = "Điểm không hợp lệ")]
        public int MultipleChoiceInCorrect { get; set; } = 0;

        [Range(0, 100, ErrorMessage = "Điểm không hợp lệ")]
        public int MultipleChoiceFail { get; set; } = 0;

        [Range(0, 100, ErrorMessage = "Điểm không hợp lệ")]
        public int EssayCorrect { get; set; } = 0;

        [Range(0, 100, ErrorMessage = "Điểm không hợp lệ")]
        public int EssayInCorrect { get; set; } = 0;

        [Range(0, 100, ErrorMessage = "Điểm không hợp lệ")]
        public int EssayFail { get; set; } = 0;

        public string? EssayResultPDF { get; set; }

        public string? Note { get; set; }

        public long ExamTrialRunId { get; set; }

        public string? CreatedDate { get; set; }

        public string? UpdatedDate { get; set; }
    }
}
