using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using Web_QM.Models;

namespace Web_QM.Controllers
{
    public class EmployeeSalaryController : Controller
    {
        private readonly QMContext _context;

        public EmployeeSalaryController(QMContext context)
        {
            _context = context;
        }

        [Authorize]
        public async Task<IActionResult> Index()
        {
            return View();
        }

        [Authorize]
        public async Task<IActionResult> GetSalaryData(string monthYear)
        {
            var employeeCode = User.FindFirstValue("EmployeeCode");

            if (string.IsNullOrEmpty(employeeCode))
            {
                return Unauthorized(new { success = false, message = "Phiên đăng nhập không hợp lệ hoặc không có mã nhân viên." });
            }

            if (string.IsNullOrEmpty(monthYear) || monthYear.Length != 7)
            {
                return BadRequest(new { success = false, message = "Định dạng Tháng/Năm không hợp lệ." });
            }

            var data = await _context.MonthlyPayroll.FirstOrDefaultAsync(q => q.EmployeeCode == employeeCode && q.DateMonth == monthYear && q.Status != 0);

            if (data == null)
            {
                return NotFound(new { success = false, message = $"Không tìm thấy bảng lương tháng {monthYear} của bạn." });
            }

            var res = new
            {
                Id = data.Id,
                EmployeeCode = data.EmployeeCode,
                DateMonth = data.DateMonth,

                ResponsibleDays = data.ResponsibleDays,
                RegularHours = data.RegularHours,
                RegularOvertimeHours = data.RegularOvertimeHours,
                SundayOvertimeHours = data.SundayOvertimeHours,
                HolidayOvertimeHours = data.HolidayOvertimeHours,
                LateNightOvertimeHours = data.LateNightOvertimeHours,
                Shift3Hours = data.Shift3Hours,
                Shift3OvertimeHours = data.Shift3OvertimeHours,
                Shift3SundayOvertimeHours = data.Shift3SundayOvertimeHours,
                BusinessTripAndPhoneFee = data.BusinessTripAndPhoneFee,
                Penalty5S = data.Penalty5S,
                UnionFee = data.UnionFee,
                PIT = data.PIT,
                PaidLeaveDay = data.PaidLeaveDay,
                AdvanceSalary = data.AdvanceSalary,
                TotalSalary = data.TotalSalary,
                Status = data.Status,

                BaseSalary = data.BaseSalary,
                InsuranceSalary = data.InsuranceSalary,
                MealAllowance = data.MealAllowance,
                DailyResponsibilityPay = data.DailyResponsibilityPay,
                FuelAllowance = data.FuelAllowance,
                ExternalWorkAllowance = data.ExternalWorkAllowance,
                HousingSubsidy = data.HousingSubsidy,
                DiligencePay = data.DiligencePay,
                NoViolationBonus = data.NoViolationBonus,
                HazardousAllowance = data.HazardousAllowance,
                CNCStressAllowance = data.CNCStressAllowance,
                SeniorityAllowance = data.SeniorityAllowance,
                CertificateAllowance = data.CertificateAllowance,
                WorkingEnvironmentAllowance = data.WorkingEnvironmentAllowance,
                JobPositionAllowance = data.JobPositionAllowance,
                MachineTestRunAllowance = data.MachineTestRunAllowance,
                RVFMachineMeasurementAllowance = data.RVFMachineMeasurementAllowance,
                StainlessSteelCleaningAllowance = data.StainlessSteelCleaningAllowance,
                FactoryGuardAllowance = data.FactoryGuardAllowance,
                HeavyDutyAllowance = data.HeavyDutyAllowance
            };

            return Ok(new { success = true, data = res });
        }

        [Authorize]
        [HttpPost]
        public async Task<IActionResult> ComplaintSalary(long monthlyPayId, string compaintContent)
        {
            if (monthlyPayId == null || string.IsNullOrEmpty(compaintContent))
            {
                return BadRequest(new { success = false, message = "Dữ liệu không hợp lệ!" });
            }
            var monthlyPay =  await _context.MonthlyPayroll.AsNoTracking().FirstOrDefaultAsync(m => m.Id == monthlyPayId);
            if (monthlyPay == null)
            {
                return BadRequest(new { success = false, message = "Bảng lương không tồn tại!" });
            }
            try
            {
                var complaint = new ComplaintSalary
                {
                    MonthlyPayId = monthlyPayId,
                    CompaintContent = compaintContent,
                    CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy")
                };
                _context.ComplaintSalary.Add(complaint);

                monthlyPay.Status = 3;
                _context.MonthlyPayroll.Update(monthlyPay);

                await _context.SaveChangesAsync();
                return Ok(new { success = true, message = "Đã gửi khiếu nại thành công!" });
            }
            catch(DbUpdateConcurrencyException)
            {
                return BadRequest(new { success = false, message = "Đã xảy ra lỗi khi gửi phản ánh. Vui lòng thử lại!" });
            }
        }

        [Authorize]
        [HttpPost]
        public async Task<IActionResult> ConfirmSalarySlip(long monthlyPayId)
        {

            if (monthlyPayId == null)
            {
                return BadRequest(new { success = false, message = "Dữ liệu không hợp lệ!" });
            }
            var monthlyPay = await _context.MonthlyPayroll.AsNoTracking().FirstOrDefaultAsync(m => m.Id == monthlyPayId);
            if (monthlyPay == null)
            {
                return BadRequest(new { success = false, message = "Bảng lương không tồn tại!" });
            }
            try
            {
                monthlyPay.Status = 2;
                _context.MonthlyPayroll.Update(monthlyPay);
                await _context.SaveChangesAsync();
                return Ok(new { success = true, message = "Đã xác nhận thành công!" });
            }
            catch (DbUpdateConcurrencyException)
            {
                return BadRequest(new { success = false, message = "Đã xảy ra lỗi khi xác nhận. Vui lòng thử lại!" });
            }
        }
    }
}
