using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models;

public partial class Feedback
{
    public long Id { get; set; }
    [Required]
    public string FeedbackerName { get; set; }
    [Required]
    public long EmployeeId { get; set; }
    [Required]
    public string Comment { get; set; }

    public string? CreatedDate { get; set; }

    public int Status { get; set; } = 0;
}
