using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class AccountPermission
    {
        [Required]
        public long AccountId { get; set; }


        [Required]
        public long PermissionId { get; set; }

        public Account account { get; set; }

        public Permission permission { get; set; }
    }
}
