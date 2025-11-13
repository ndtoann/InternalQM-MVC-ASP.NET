using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class MachineGroup
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập nhóm máy")]
        public string GroupName {  get; set; }

        [Required(ErrorMessage = "Vui lòng nhập loại nhóm máy")]
        public string MachineType { get; set; }

        public string? Standard {  get; set; }
    }
}
