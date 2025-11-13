using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class Permission
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập giá trị")]
        public string ClaimValue { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập mô tả")]
        public string Description { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập phân luồng")]
        public string Module { get; set; }
    }
}
