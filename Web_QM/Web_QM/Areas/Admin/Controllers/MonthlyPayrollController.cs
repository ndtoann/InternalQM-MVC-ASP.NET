using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using System.Linq;
using Web_QM.Models;
using Web_QM.Models.ViewModels;

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class MonthlyPayrollController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public MonthlyPayrollController(QMContext context, IWebHostEnvironment env  )
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewMonthlyPayroll")]
        public async Task<IActionResult> Index(int? status, string department, string key, string month)
        {
            var salaryQuery = _context.MonthlyPayroll.AsNoTracking();
            var employeeQuery = _context.Employees.AsNoTracking();

            if (status.HasValue)
            {
                salaryQuery = salaryQuery.Where(s => s.Status == status);
            }

            var resultQuery = from salary in salaryQuery
                              join employee in employeeQuery
                              on salary.EmployeeCode equals employee.EmployeeCode
                              select new
                              {
                                  Id = salary.Id,
                                  RegularHours = salary.RegularHours,
                                  ResponsibleDays = salary.ResponsibleDays,
                                  RegularOvertimeHours = salary.RegularOvertimeHours,
                                  SundayOvertimeHours = salary.SundayOvertimeHours,
                                  HolidayOvertimeHours = salary.HolidayOvertimeHours,
                                  LateNightOvertimeHours = salary.LateNightOvertimeHours,
                                  TotalSalary = salary.TotalSalary,
                                  EmployeeCode = employee.EmployeeCode,
                                  EmployeeName = employee.EmployeeName,
                                  Department = employee.Department,
                                  DateMonth = salary.DateMonth,
                                  Status = salary.Status
                              };

            if (!string.IsNullOrEmpty(department))
            {
                resultQuery = resultQuery.Where(item =>
                    item.Department == department
                );
            }

            if (!string.IsNullOrEmpty(month))
            {
                resultQuery = resultQuery.Where(item =>
                    item.DateMonth == month
                );
            }

            if (!string.IsNullOrEmpty(key))
            {
                string searchKey = key.ToLower();
                resultQuery = resultQuery.Where(item =>
                    item.EmployeeCode.ToLower().Contains(searchKey) ||
                    item.EmployeeName.ToLower().Contains(searchKey)
                );
            }

            var res = await resultQuery.ToListAsync();

            var viewModelList = res.Select(item => new MonthlyPayrollView
            {
                Id = item.Id,
                RegularHours = item.RegularHours,
                ResponsibleDays = item.ResponsibleDays,
                RegularOvertimeHours = item.RegularOvertimeHours ?? 0,
                SundayOvertimeHours = item.SundayOvertimeHours ?? 0,
                HolidayOvertimeHours = item.HolidayOvertimeHours ?? 0,
                LateNightOvertimeHours = item.LateNightOvertimeHours ?? 0,
                TotalSalary = item.TotalSalary ?? 0,
                EmployeeCode = item.EmployeeCode,
                EmployeeName = item.EmployeeName,
                Department = item.Department,
                DateMonth = item.DateMonth,
                Status = item.Status ?? 0
            }).OrderByDescending(o => o.DateMonth).Take(500).ToList();

            var statusOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "", Text = "Trạng thái phiếu" },
                new SelectListItem { Value = "0", Text = "Nháp" },
                new SelectListItem { Value = "1", Text = "Chờ xác nhận" },
                new SelectListItem { Value = "2", Text = "Đã xác nhận" },
                new SelectListItem { Value = "3", Text = "Khiếu nại" }
            };
            ViewData["Status"] = new SelectList(statusOptions, "Value", "Text", status);

            var departments = await _context.Departments.AsNoTracking().ToListAsync();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName,
                Selected = d.DepartmentName == department
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewData["Departments"] = departmentsList;

            ViewBag.MonthSearch = month;
            ViewBag.KeySearch = key;

            return View(viewModelList);
        }

        [Authorize(Policy = "ViewMonthlyPayroll")]
        public async Task<IActionResult> GetDetails(long id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var salaryDetails = await _context.MonthlyPayroll.AsNoTracking().FirstOrDefaultAsync(s => s.Id == id);
            if (salaryDetails == null)
            {
                return NotFound();
            }
            return Json(salaryDetails);
        }

        [Authorize(Policy = "UpdateMonthlyPayroll")]
        public async Task<IActionResult> Add()
        {
            var employees = await _context.Employees.AsNoTracking().OrderBy(o => o.EmployeeCode).ToListAsync();
            var employeeList = employees.Select(e => new
            {
                Value = e.EmployeeCode,
                Text = e.EmployeeCode + " - " + e.EmployeeName
            });
            ViewBag.ListEmployee = new SelectList(employeeList, "Value", "Text");

            return View();
        }

        [Authorize(Policy = "UpdateMonthlyPayroll")]
        [HttpPost]
        public async Task<IActionResult> Add(MonthlyPayroll monthlyPayroll)
        {
            var employees = await _context.Employees.AsNoTracking().OrderBy(o => o.EmployeeCode).ToListAsync();
            var employeeList = employees.Select(e => new
            {
                Value = e.EmployeeCode,
                Text = e.EmployeeCode + " - " + e.EmployeeName
            });
            ViewBag.ListEmployee = new SelectList(employeeList, "Value", "Text");

            if (!ModelState.IsValid)
            {
                return View(monthlyPayroll);
            }

            var empl = await _context.Employees.AsNoTracking().FirstOrDefaultAsync(e => e.EmployeeCode == monthlyPayroll.EmployeeCode);
            if (empl == null)
            {
                ModelState.AddModelError("EmployeeCode", "Mã nhân viên không hợp lệ!");
                return View(monthlyPayroll);
            }

            var emplSalary = await _context.Salaries.AsNoTracking().FirstOrDefaultAsync(s => s.EmployeeCode == monthlyPayroll.EmployeeCode);
            if (emplSalary == null) {
                ModelState.AddModelError("EmployeeCode", "Chưa có thông tin lương cho nhân viên này!");
                return View(monthlyPayroll);
            }

            bool isDuplicate = await IsDuplicate(monthlyPayroll.EmployeeCode, monthlyPayroll.DateMonth);
            if (isDuplicate)
            {
                ModelState.AddModelError("EmployeeCode", "Nhân viên đã có thông tin lương tháng này!");
                return View(monthlyPayroll);
            }
            try
            {
                monthlyPayroll.RegularOvertimeHours = monthlyPayroll.RegularOvertimeHours ?? 0;
                monthlyPayroll.SundayOvertimeHours = monthlyPayroll.SundayOvertimeHours ?? 0;
                monthlyPayroll.HolidayOvertimeHours = monthlyPayroll.HolidayOvertimeHours ?? 0;
                monthlyPayroll.LateNightOvertimeHours = monthlyPayroll.LateNightOvertimeHours ?? 0;
                monthlyPayroll.Shift3Hours = monthlyPayroll.Shift3Hours ?? 0;
                monthlyPayroll.Shift3OvertimeHours = monthlyPayroll.Shift3OvertimeHours ?? 0;
                monthlyPayroll.Shift3SundayOvertimeHours = monthlyPayroll.Shift3SundayOvertimeHours ?? 0;
                monthlyPayroll.BusinessTripAndPhoneFee = monthlyPayroll.BusinessTripAndPhoneFee ?? 0;
                monthlyPayroll.Penalty5S = monthlyPayroll.Penalty5S ?? 0;
                monthlyPayroll.UnionFee = monthlyPayroll.UnionFee ?? 0;
                monthlyPayroll.PIT = monthlyPayroll.PIT ?? 0;
                monthlyPayroll.PaidLeaveDay = monthlyPayroll.PaidLeaveDay ?? 0;
                monthlyPayroll.AdvanceSalary = monthlyPayroll.AdvanceSalary ?? 0;
                monthlyPayroll.Status = 0;
                monthlyPayroll.TotalSalary = CalculateTotalSalary(emplSalary, monthlyPayroll);
                monthlyPayroll.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");

                //lương phụ cấp hiện tại lấy theo bảng lương nhân viên
                monthlyPayroll.LevelSalary = emplSalary.LevelSalary;
                monthlyPayroll.BaseSalary = emplSalary.BaseSalary;
                monthlyPayroll.InsuranceSalary = emplSalary.InsuranceSalary;
                monthlyPayroll.MealAllowance = emplSalary.MealAllowance;
                monthlyPayroll.DailyResponsibilityPay = emplSalary.DailyResponsibilityPay;
                monthlyPayroll.FuelAllowance = emplSalary.FuelAllowance;
                monthlyPayroll.ExternalWorkAllowance = emplSalary.ExternalWorkAllowance;
                monthlyPayroll.HousingSubsidy = emplSalary.HousingSubsidy;
                monthlyPayroll.DiligencePay = emplSalary.DiligencePay;
                monthlyPayroll.NoViolationBonus = emplSalary.NoViolationBonus;
                monthlyPayroll.HazardousAllowance = emplSalary.HazardousAllowance;
                monthlyPayroll.CNCStressAllowance = emplSalary.CNCStressAllowance;
                monthlyPayroll.SeniorityAllowance = emplSalary.SeniorityAllowance;
                monthlyPayroll.CertificateAllowance = emplSalary.CertificateAllowance;
                monthlyPayroll.WorkingEnvironmentAllowance = emplSalary.WorkingEnvironmentAllowance;
                monthlyPayroll.JobPositionAllowance = emplSalary.JobPositionAllowance;
                monthlyPayroll.MachineTestRunAllowance = emplSalary.MachineTestRunAllowance;
                monthlyPayroll.RVFMachineMeasurementAllowance = emplSalary.RVFMachineMeasurementAllowance;
                monthlyPayroll.StainlessSteelCleaningAllowance = emplSalary.StainlessSteelCleaningAllowance;
                monthlyPayroll.FactoryGuardAllowance = emplSalary.FactoryGuardAllowance;
                monthlyPayroll.HeavyDutyAllowance = emplSalary.HeavyDutyAllowance;

                _context.MonthlyPayroll.Add(monthlyPayroll);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(monthlyPayroll);
            }
        }

        [Authorize(Policy = "UpdateMonthlyPayroll")]
        public async Task<IActionResult> Edit(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var salaryToEdit = await _context.MonthlyPayroll.AsNoTracking().FirstOrDefaultAsync(s => s.Id == id);
            if (salaryToEdit == null)
            {
                return NotFound();
            }

            if(salaryToEdit.Status == 1 || salaryToEdit.Status == 2)
            {
                TempData["ErrorMessage"] = "Không thể chỉnh sửa phiếu lương!";
                return RedirectToAction(nameof(Index));
            }

            var employees = await _context.Employees.AsNoTracking().OrderBy(o => o.EmployeeCode).ToListAsync();
            var employeeList = employees.Select(e => new
            {
                Value = e.EmployeeCode,
                Text = e.EmployeeCode + " - " + e.EmployeeName
            });
            ViewBag.ListEmployee = new SelectList(employeeList, "Value", "Text");

            return View(salaryToEdit);
        }

        [Authorize(Policy = "UpdateMonthlyPayroll")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, MonthlyPayroll monthlyPayroll)
        {
            if (id != monthlyPayroll.Id)
            {
                return NotFound();
            }
            var employees = await _context.Employees.AsNoTracking().OrderBy(o => o.EmployeeCode).ToListAsync();
            var employeeList = employees.Select(e => new
            {
                Value = e.EmployeeCode,
                Text = e.EmployeeCode + " - " + e.EmployeeName
            });
            ViewBag.ListEmployee = new SelectList(employeeList, "Value", "Text");
            if (!ModelState.IsValid)
            {
                return View(monthlyPayroll);
            }

            var oldSalary = await _context.MonthlyPayroll.AsNoTracking().FirstOrDefaultAsync(s => s.Id == monthlyPayroll.Id);
            if (oldSalary == null)
            {
                return NotFound();
            }

            if (oldSalary.Status == 1 || oldSalary.Status == 2)
            {
                TempData["ErrorMessage"] = "Không thể chỉnh sửa phiếu lương!";
                return RedirectToAction(nameof(Index));
            }

            var empl = await _context.Employees.AsNoTracking().FirstOrDefaultAsync(e => e.EmployeeCode == monthlyPayroll.EmployeeCode);
            if (empl == null)
            {
                ModelState.AddModelError("EmployeeCode", "Mã nhân viên không hợp lệ!");
                return View(monthlyPayroll);
            }

            var emplSalary = await _context.Salaries.AsNoTracking().FirstOrDefaultAsync(s => s.EmployeeCode == monthlyPayroll.EmployeeCode);
            if (emplSalary == null)
            {
                ModelState.AddModelError("EmployeeCode", "Chưa có thông tin lương cho nhân viên này!");
                return View(monthlyPayroll);
            }

            bool isDuplicate = await IsDuplicate(monthlyPayroll.EmployeeCode, monthlyPayroll.DateMonth, monthlyPayroll.Id);
            if (isDuplicate)
            {
                ModelState.AddModelError("EmployeeCode", "Nhân viên đã có thông tin lương!");
                return View(monthlyPayroll);
            }
            try
            {
                monthlyPayroll.RegularOvertimeHours = monthlyPayroll.RegularOvertimeHours ?? 0;
                monthlyPayroll.SundayOvertimeHours = monthlyPayroll.SundayOvertimeHours ?? 0;
                monthlyPayroll.HolidayOvertimeHours = monthlyPayroll.HolidayOvertimeHours ?? 0;
                monthlyPayroll.LateNightOvertimeHours = monthlyPayroll.LateNightOvertimeHours ?? 0;
                monthlyPayroll.Shift3Hours = monthlyPayroll.Shift3Hours ?? 0;
                monthlyPayroll.Shift3OvertimeHours = monthlyPayroll.Shift3OvertimeHours ?? 0;
                monthlyPayroll.Shift3SundayOvertimeHours = monthlyPayroll.Shift3SundayOvertimeHours ?? 0;
                monthlyPayroll.BusinessTripAndPhoneFee = monthlyPayroll.BusinessTripAndPhoneFee ?? 0;
                monthlyPayroll.Penalty5S = monthlyPayroll.Penalty5S ?? 0;
                monthlyPayroll.UnionFee = monthlyPayroll.UnionFee ?? 0;
                monthlyPayroll.PIT = monthlyPayroll.PIT ?? 0;
                monthlyPayroll.PaidLeaveDay = monthlyPayroll.PaidLeaveDay ?? 0;
                monthlyPayroll.AdvanceSalary = monthlyPayroll.AdvanceSalary ?? 0;
                monthlyPayroll.Status = 0;
                monthlyPayroll.TotalSalary = CalculateTotalSalary(emplSalary, monthlyPayroll);
                monthlyPayroll.CreatedDate = oldSalary.CreatedDate;
                monthlyPayroll.UpdatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");

                //lương phụ cấp hiện tại lấy theo bảng lương nhân viên
                monthlyPayroll.LevelSalary = emplSalary.LevelSalary;
                monthlyPayroll.BaseSalary = emplSalary.BaseSalary;
                monthlyPayroll.InsuranceSalary = emplSalary.InsuranceSalary;
                monthlyPayroll.MealAllowance = emplSalary.MealAllowance;
                monthlyPayroll.DailyResponsibilityPay = emplSalary.DailyResponsibilityPay;
                monthlyPayroll.FuelAllowance = emplSalary.FuelAllowance;
                monthlyPayroll.ExternalWorkAllowance = emplSalary.ExternalWorkAllowance;
                monthlyPayroll.HousingSubsidy = emplSalary.HousingSubsidy;
                monthlyPayroll.DiligencePay = emplSalary.DiligencePay;
                monthlyPayroll.NoViolationBonus = emplSalary.NoViolationBonus;
                monthlyPayroll.HazardousAllowance = emplSalary.HazardousAllowance;
                monthlyPayroll.CNCStressAllowance = emplSalary.CNCStressAllowance;
                monthlyPayroll.SeniorityAllowance = emplSalary.SeniorityAllowance;
                monthlyPayroll.CertificateAllowance = emplSalary.CertificateAllowance;
                monthlyPayroll.WorkingEnvironmentAllowance = emplSalary.WorkingEnvironmentAllowance;
                monthlyPayroll.JobPositionAllowance = emplSalary.JobPositionAllowance;
                monthlyPayroll.MachineTestRunAllowance = emplSalary.MachineTestRunAllowance;
                monthlyPayroll.RVFMachineMeasurementAllowance = emplSalary.RVFMachineMeasurementAllowance;
                monthlyPayroll.StainlessSteelCleaningAllowance = emplSalary.StainlessSteelCleaningAllowance;
                monthlyPayroll.FactoryGuardAllowance = emplSalary.FactoryGuardAllowance;
                monthlyPayroll.HeavyDutyAllowance = emplSalary.HeavyDutyAllowance;

                _context.MonthlyPayroll.Update(monthlyPayroll);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(monthlyPayroll);
            }
        }

        [Authorize(Policy = "UpdateMonthlyPayroll")]
        public async Task<IActionResult> Delete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var salaryToDelete = await _context.MonthlyPayroll.FindAsync(id);
            if (salaryToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy dữ liệu!";
                return RedirectToAction(nameof(Index));
            }
            try
            {
                await _context.ComplaintSalary
                     .Where(a => a.MonthlyPayId == salaryToDelete.Id)
                     .ExecuteDeleteAsync();

                _context.MonthlyPayroll.Remove(salaryToDelete);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa dữ liệu!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Index));
        }

        [Authorize(Policy = "UpdateMonthlyPayroll")]
        [HttpPost]
        public async Task<IActionResult> ApproveVotes([FromBody] List<long> ids)
        {
            if (ids == null || !ids.Any())
            {
                return BadRequest(new { success = false, message = "Không có ID nào được chọn để duyệt." });
            }

            try
            {
                var payrollsToUpdate = await _context.MonthlyPayroll
                                                     .Where(p => ids.Contains(p.Id))
                                                     .ToListAsync();

                if (!payrollsToUpdate.Any())
                {
                    return NotFound(new { success = false, message = "Không tìm thấy bản ghi lương nào với các ID đã chọn." });
                }

                int updatedCount = 0;
                foreach (var payroll in payrollsToUpdate)
                {
                    if (payroll.Status != 2 && payroll.Status != 1)
                    {
                        payroll.Status = 1;
                        updatedCount++;
                    }
                }

                if (updatedCount > 0)
                {
                    await _context.SaveChangesAsync();
                }

                return Ok(new
                {
                    success = true,
                    message = "Đã duyệt thành công!"
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = "Lỗi máy chủ: Không thể duyệt dữ liệu." });
            }
        }

        private async Task<bool> IsDuplicate(string employeeCode, string dateMonth, long currentId = 0)
        {
            bool exists = await _context.MonthlyPayroll
                .AnyAsync(p =>
                    p.EmployeeCode == employeeCode &&
                    p.DateMonth == dateMonth &&
                    p.Id != currentId
                );
            return exists;
        }

        [Authorize(Policy = "ViewMonthlyPayroll")]
        public async Task<IActionResult> GetComplaintHistory(long monthlyPayId)
        {
            if (monthlyPayId == null)
            {
                return BadRequest(new { message = "Bảng lương không hợp lệ." });
            }

            try
            {
                var complaints = await _context.ComplaintSalary
                                               .Where(c => c.MonthlyPayId == monthlyPayId)
                                               .OrderByDescending(c => c.Id)
                                               .ToListAsync();
                return Json(complaints);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Lỗi hệ thống khi truy vấn dữ liệu: " + ex.Message });
            }
        }

        private static decimal CalculateTotalSalary(Salary emplSalary, MonthlyPayroll monthlyPayroll)
        {
            const decimal REGULAR_OT_FACTOR = 1.5m; //tăng ca thường
            const decimal SUNDAY_OT_FACTOR = 2.0m; //tăng ca chủ nhật
            const decimal HOLIDAY_OT_FACTOR = 3.0m; //tăng ca ngày lễ
            const decimal LATE_NIGHT_OT_FACTOR = 1.3m; //tăng ca đêm
            const decimal WORKING_DAYS_IN_MONTH = 25m; // số ngày làm việc quy định
            const decimal WORKING_HOURS_IN_DAY = 8m;   // số giờ làm việc quy định
            const decimal INSURANCE = 0.105m;   // bảo hiểm 10.5%

            decimal dailyBasePay = emplSalary.BaseSalary / WORKING_DAYS_IN_MONTH;

            decimal hourlyBasePay = dailyBasePay / WORKING_HOURS_IN_DAY;

            decimal dailyResponsibilityPay = emplSalary.DailyResponsibilityPay ?? 0;
            dailyResponsibilityPay = dailyResponsibilityPay * monthlyPayroll.ResponsibleDays;
            decimal totalAllowance = dailyResponsibilityPay +
                                     (emplSalary.FuelAllowance ?? 0) +
                                     (emplSalary.ExternalWorkAllowance ?? 0) +
                                     (emplSalary.HousingSubsidy ?? 0) +
                                     (emplSalary.DiligencePay ?? 0) +
                                     (emplSalary.NoViolationBonus ?? 0) +
                                     (emplSalary.HazardousAllowance ?? 0) +
                                     (emplSalary.CNCStressAllowance ?? 0) +
                                     (emplSalary.SeniorityAllowance ?? 0) +
                                     (emplSalary.CertificateAllowance ?? 0) +
                                     (emplSalary.WorkingEnvironmentAllowance ?? 0) +
                                     (emplSalary.JobPositionAllowance ?? 0) +
                                     (emplSalary.MachineTestRunAllowance ?? 0) +
                                     (emplSalary.RVFMachineMeasurementAllowance ?? 0) +
                                     (emplSalary.StainlessSteelCleaningAllowance ?? 0) +
                                     (emplSalary.FactoryGuardAllowance ?? 0) +
                                     (emplSalary.HeavyDutyAllowance ?? 0);

            decimal regularPay = hourlyBasePay * monthlyPayroll.RegularHours;

            decimal overtimePay =
                hourlyBasePay * (monthlyPayroll.RegularOvertimeHours ?? 0) * REGULAR_OT_FACTOR +
                hourlyBasePay * (monthlyPayroll.SundayOvertimeHours ?? 0) * SUNDAY_OT_FACTOR +
                hourlyBasePay * (monthlyPayroll.HolidayOvertimeHours ?? 0) * HOLIDAY_OT_FACTOR +
                hourlyBasePay * (monthlyPayroll.LateNightOvertimeHours ?? 0) * LATE_NIGHT_OT_FACTOR;

            decimal shift3Pay = 0;

            decimal nonHourlyAllowance = monthlyPayroll.BusinessTripAndPhoneFee ?? 0;

            decimal grossSalary = regularPay + overtimePay + totalAllowance + nonHourlyAllowance + shift3Pay;

            // trừ lương
            decimal insuranceDeduction = (emplSalary.InsuranceSalary ?? 0) * INSURANCE;

            decimal totalDeductions = (monthlyPayroll.Penalty5S ?? 0) +
                                      (monthlyPayroll.UnionFee ?? 0) +
                                      (monthlyPayroll.PIT ?? 0) +
                                      insuranceDeduction;

            decimal finalTotalSalary = grossSalary - totalDeductions - (monthlyPayroll.AdvanceSalary ?? 0);

            return finalTotalSalary;
        }
    }
}
