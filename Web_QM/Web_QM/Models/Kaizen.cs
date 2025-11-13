using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class Kaizen
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập ngày/tháng")]
        public DateOnly DateMonth { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập mã nhân viên")]
        public string EmployeeCode { get; set; }

        public string EmployeeName { get; set; }

        public string Department {  get; set; }

        public string? AppliedDepartment { get; set; }

        public string? ImprovementGoal { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập tiêu đề cải tiến")]
        public string ImprovementTitle { get; set; }

        public string? CurrentSituation { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập ý kiến cài tiến")]
        public string ProposedIdea { get; set; }

        public string? EstimatedBenefit { get; set; }

        public string? TeamLeaderRating { get; set; }

        public string? ManagementReview { get; set; }

        public string? Picture { get; set; }

        public string? Deadline { get; set; }

        public string? StartTime { get; set; }

        public string? CurrentStatus { get; set; }

        public string? Note { get; set; }
    }
}
