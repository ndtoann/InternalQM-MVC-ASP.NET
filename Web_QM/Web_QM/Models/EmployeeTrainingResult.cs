using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models;

public partial class EmployeeTrainingResult
{
    public long Id { get; set; }
    [Required]
    public long EmployeeId { get; set; }
    [Required]
    public long TrainingId { get; set; }
    [Required]
    public int Status { get; set; }
    [Required]
    public string EvaluationPeriod { get; set; }
    [Required]
    public int Score { get; set; }
}
