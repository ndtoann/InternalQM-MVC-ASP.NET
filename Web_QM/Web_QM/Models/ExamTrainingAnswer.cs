using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class ExamTrainingAnswer
    {
        public long Id { get; set; }
        [Required]
        public string EmployeeName { get; set; }
        [Required]
        public string EmployeeCode { get; set; }
        [Required]
        public string ListAnswer { get; set; }

        public int? IsShow { get; set; }
        [Range(0, 100, ErrorMessage = "Điểm không hợp lệ")]
        public int? TnPoint { get; set; }
        [Range(0, 100, ErrorMessage = "Điểm không hợp lệ")]
        public int? TlPoint { get; set; }

        public string? EssayResultPDF { get; set; }

        public string? Note { get; set; }

        public long ExamTrainingId { get; set; }

        public string? CreatedDate { get; set; }

        public string? UpdatedDate { get; set; }
    }
}
