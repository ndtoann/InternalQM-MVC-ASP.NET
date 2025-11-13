using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models;

public partial class Training
{
    public long Id { get; set; }
    [Required(ErrorMessage = "Nội dung đào tạo không được để trống")]
    public string TrainingName { get; set; }
    [Required(ErrorMessage = "Phân loại đào tạo không được để trống")]
    public string Type { get; set; }

    public string? CreatedDate { get; set; }

    public string? UpdatedDate { get; set; }
}
