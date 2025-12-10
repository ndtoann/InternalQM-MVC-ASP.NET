using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Linq;
using Web_QM.Models;
using Web_QM.Models.ViewModels;

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class TimeSheetController : Controller
    {
        private readonly QMContext _context;

        public TimeSheetController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ViewTimeSheet")]
        public async Task<IActionResult> Index(string dateMonth, string key)
        {
            ViewBag.DefaultMonth = dateMonth;
            ViewBag.KeySearch = key;

            var permissions = User.Claims.Where(c => c.Type == "Permission").Select(c => c.Value).ToList();
            var canViewAll = permissions.Any(p => p.Equals("TimeSheet.ViewAll", StringComparison.OrdinalIgnoreCase));
            var canApprove = permissions.Any(p => p.Equals("TimeSheet.Approve", StringComparison.OrdinalIgnoreCase));
            var userDepartment = User.FindFirst("Department")?.Value;

            var totalHoursByTimesheet = await _context.Timekeepings
                .GroupBy(t => t.TimesheetId)
                .Select(g => new
                {
                    TimesheetId = g.Key,
                    TotalHours = g.Sum(t => t.TotalHours) ?? 0M
                })
                .ToDictionaryAsync(x => x.TimesheetId, x => x.TotalHours);

            var timesheets = _context.Timesheets
                .Where(t => t.Status != 1)
                .Join(
                    _context.Employees,
                    timesheet => timesheet.EmployeeId,
                    employee => employee.Id,
                    (timesheet, employee) => new { Timesheet = timesheet, Employee = employee }
                );

            if (!canViewAll && !string.IsNullOrEmpty(userDepartment))
            {
                timesheets = timesheets.Where(t => t.Employee.Department == userDepartment);
            }
            if (!canApprove)
            {
                timesheets = timesheets.Where(t => t.Timesheet.Status == 3);
            }

            if (!string.IsNullOrEmpty(dateMonth))
            {
                timesheets = timesheets.Where(t => t.Timesheet.DateMonth == dateMonth);
            }

            if (!string.IsNullOrEmpty(key))
            {
                string searchKey = key.Trim().ToLower();
                timesheets = timesheets.Where(t =>
                    t.Employee.EmployeeCode.ToLower().Contains(searchKey) ||
                    t.Employee.EmployeeName.ToLower().Contains(searchKey)
                );
            }

            var results = await timesheets
                .Select(x => new TimeSheetView
                {
                    Id = x.Timesheet.Id,
                    EmployeeId = x.Employee.Id,
                    EmployeeCode = x.Employee.EmployeeCode,
                    EmployeeName = x.Employee.EmployeeName,
                    DateMonth = x.Timesheet.DateMonth,
                    TotalHours = totalHoursByTimesheet.ContainsKey(x.Timesheet.Id)
                    ? totalHoursByTimesheet[x.Timesheet.Id]
                    : 0M,
                    Status = x.Timesheet.Status
                })
                .ToListAsync();

            return View(results);
        }

        [Authorize(Policy = "ViewTimeSheet")]
        public async Task<IActionResult> GetTimekeeping(long timesheetId)
        {
            var timekeepingData = await _context.Timekeepings
                .Where(d => d.TimesheetId == timesheetId)
                .OrderBy(d => d.WorkDate)
                .ToListAsync();

            var data = timekeepingData
                .Select(d => new
                {
                    workDate = d.WorkDate.ToString("yyyy-MM-dd"),
                    shift = d.Shift,
                    note = d.Note,
                    timeIn = d.TimeIn.HasValue ? d.TimeIn.Value.ToString("c") : null,
                    timeOut = d.TimeOut.HasValue ? d.TimeOut.Value.ToString("c") : null,
                    totalHours = d.TotalHours
                })
                .ToList();

            return Json(data);
        }

        [Authorize(Policy = "ApproveTimeSheet")]
        [HttpPost]
        public async Task<IActionResult> ApproveMultiple([FromBody] BulkApproveModel model)
        {
            if (model == null || model.Ids == null || !model.Ids.Any())
            {
                return Json(new { success = false, message = "Không có phiếu nào được chọn." });
            }

            var timesheetsToApprove = await _context.Timesheets
                .Where(t => model.Ids.Contains(t.Id) && t.Status == 2)
                .ToListAsync();

            if (!timesheetsToApprove.Any())
            {
                return Json(new { success = false, message = "Không tìm thấy phiếu chờ duyệt hợp lệ nào." });
            }

            foreach (var sheet in timesheetsToApprove)
            {
                sheet.Status = 3;
            }

            await _context.SaveChangesAsync();

            return Json(new { success = true, count = timesheetsToApprove.Count });
        }

        [Authorize(Policy = "ApproveTimeSheet")]
        [HttpPost]
        public async Task<IActionResult> RejectMultiple([FromBody] BulkApproveModel model)
        {
            if (model == null || model.Ids == null || !model.Ids.Any())
            {
                return Json(new { success = false, message = "Không có phiếu nào được chọn." });
            }

            var timesheetsToReject = await _context.Timesheets
                .Where(t => model.Ids.Contains(t.Id) && t.Status == 2)
                .ToListAsync();

            if (!timesheetsToReject.Any())
            {
                return Json(new { success = false, message = "Không tìm thấy phiếu chờ duyệt hợp lệ nào." });
            }

            foreach (var sheet in timesheetsToReject)
            {
                sheet.Status = 4;
            }

            await _context.SaveChangesAsync();

            return Json(new { success = true, count = timesheetsToReject.Count });
        }

        public class BulkApproveModel
        {
            public List<long> Ids { get; set; }
        }
    }
}
