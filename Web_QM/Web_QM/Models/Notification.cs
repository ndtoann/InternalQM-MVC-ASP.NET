using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class Notification
    {
        public long Id { get; set; }
        [Required]
        public string Message { get; set; }
        [Required]
        public int IsRead { get; set; }

        public string? CreatedDate { get; set; }
    }
}
