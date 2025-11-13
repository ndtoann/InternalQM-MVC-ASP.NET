using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class ErrorData
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập số order")]
        public string OrderNo { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập tên chi tiết")]
        public string PartName { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập số lượng order")]
        [Range(0,10000)]
        public int QtyOrder { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập số lượng NG")]
        [Range(0, 10000)]
        public int QtyNG { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập ngày/tháng")]
        public DateOnly DateMonth { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập phát hiện lỗi")]
        public string ErrorDetected { get; set; }  //phát hiện lỗi

        [Required(ErrorMessage = "Vui lòng nhập dạng lỗi")]
        public string ErrorType { get; set; }  //dạng lỗi

        public string? ErrorCause { get; set; }  //nguyên nhân lỗi

        [Required(ErrorMessage = "Vui lòng nhập nội dung lỗi")]
        public string ErrorContent { get; set; }  //nội dung

        public string? ToleranceAssessment { get; set; }  //nhận định dung sai

        public string? Reason { get; set; }  //nguyên nhân

        public string? Countermeasure { get; set; }  //đối sách

        [Required(ErrorMessage = "Vui lòng nhập NCC")]
        public string NCC { get; set; }

        public string? EmployeeCode { get; set; }

        public string? Department { get; set; }

        public DateOnly? ErrorCompletionDate { get; set; }  //ngày hoàn thành giấy báo lỗi

        public string? RemedialMeasures { get; set; }  //biện pháp khắc phục

        public string? Note { get; set; }

        public string? TimeWriteError { get; set; }  //thời gian viết giấy lỗi

        public string? ReviewNnds { get; set; }  //rà soát thực hiên NNDS
    }
}
