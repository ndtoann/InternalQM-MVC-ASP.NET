using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class EquipmentRepairHistory
    {
        public long Id { get; set; }

        public string? EquipmentCode { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập tên thiết bị")]
        public string EquipmentName { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập tình trạng lỗi hỏng")]
        public string ErrorCondition { get; set; }

        [Range(typeof(int), "1", "999999", ErrorMessage = "Giá trị nhỏ nhất là 1")]
        public int Qty { get; set; } = 1;

        public string? Department { get; set; }

        public string? Reason { get; set; }

        public string? ProcessingMethod { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập ngày tiếp nhận lỗi")]
        public DateOnly DateMonth {  get; set; }

        public DateOnly? CompletionDate { get; set; }

        public DateOnly? ConfirmCompletionDate { get; set; }

        public string? RemedialStaff { get; set; }

        public string? RecipientOfRepairedDevice { get; set; }

        public string? RepairCosts { get; set; }

        public string? Note {  get; set; }
    }
}
