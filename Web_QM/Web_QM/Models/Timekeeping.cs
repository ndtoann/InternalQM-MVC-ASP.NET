using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class Timekeeping
    {
        public long Id { get; set; }

        [Required]
        public long TimesheetId { get; set; }

        [Required]
        public DateOnly WorkDate { get; set; }

        public string? Note { get; set; }
        public TimeSpan? TimeIn { get; set; }
        public TimeSpan? TimeOut { get; set; }
        public string? Shift { get; set; }
        public decimal? TotalHours { get; set; }
    }
}
