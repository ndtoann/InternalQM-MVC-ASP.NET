using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public partial class Account
    {
        public long Id { get; set; }
        [Required(ErrorMessage = "Tài khoản không được để trống")]
        public string UserName { get; set; }
        [Required(ErrorMessage = "Mật khẩu không được để trống")]
        [RegularExpression(@"^[A-Za-z0-9@]*$", ErrorMessage = "Không thể chứa ký tự đặc biệt")]
        [MinLength(6, ErrorMessage = "Mật khẩu tối thiểu 6 ký tự")]
        public string Password { get; set; }
        [Required]
        public string Salt { get; set; }
        [Required(ErrorMessage = "Tên nhân viên không được để trống")]
        public string StaffName { get; set; }
        [Required(ErrorMessage = "Mã nhân viên không được để trống")]
        public string StaffCode { get; set; }

        public string? CreatedDate { get; set; }

        public string? UpdatedDate { get; set; }

    }
}
