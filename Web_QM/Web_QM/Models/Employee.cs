using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models;

public partial class Employee
{
    public long Id { get; set; }
    [Required(ErrorMessage = "Mã nhân viên không được để trống")]
    public string EmployeeCode { get; set; }
    [Required(ErrorMessage = "Tên nhân viên không được để trống")]
    public string EmployeeName { get; set; }

    public DateOnly? DateOfBirth { get; set; }

    public string? Gender { get; set; }
    [Required(ErrorMessage = "Vui lòng chọn bộ phận")]
    public string Department { get; set; }

    [Required(ErrorMessage = "Vui lòng nhập ngày vào công ty")]
    public DateOnly HireDate { get; set; }
    [Required(ErrorMessage = "Vui lòng chọn chức vụ")]
    public string Position { get; set; }

    public string? Avatar { get; set; }

    public string? Note { get; set; }

    public string? CreatedDate { get; set; }

    public string? UpdatedDate { get; set; }
}
