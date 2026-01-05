using DocumentFormat.OpenXml.InkML;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Web_QM.Models;

namespace Web_QM.Areas.HR.Controllers
{
    [Area("HR")]
    public class DashboardController : Controller
    {
        private readonly QMContext _context;

        public DashboardController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ViewHr")]
        public async Task<IActionResult> Index()
        {
            var cEmployee = await _context.Employees.CountAsync();
            var cKaizen = await _context.Kaizens.CountAsync();
            var cError = await _context.ErrorDatas.CountAsync();
            var c5S = await _context.EmployeeViolation5S.CountAsync();

            ViewBag.CountEmployee = cEmployee;
            ViewBag.CountKaizen = cKaizen;
            ViewBag.CountError = cError;
            ViewBag.Count5S = c5S;

            var cMachine = await _context.Machines.CountAsync();

            ViewBag.CountMachine = cMachine;

            var departmentData = await _context.Employees
                                  .GroupBy(e => e.Department)
                                  .Select(g => new {
                                      DepartmentName = g.Key,
                                      EmployeeCount = g.Count()
                                  })
                                  .ToListAsync();

            var labelPie = departmentData.Select(d => d.DepartmentName).ToList();
            var dataPie = departmentData.Select(d => d.EmployeeCount).ToList();

            ViewBag.LabelPie = labelPie;
            ViewBag.DataPie = dataPie;

            var currentYear = DateOnly.FromDateTime(DateTime.Now).Year;

            var monthlyKaizenCounts = await _context.Kaizens
                                           .AsNoTracking()
                                           .Where(e => e.DateMonth.Year == currentYear)
                                           .GroupBy(e => e.DateMonth.Month)
                                           .Select(g => new {
                                               Month = g.Key,
                                               KaizenCount = g.Count()
                                           })
                                           .ToListAsync();

            var dataBar = new List<int>();
            for (int month = 1; month <= 12; month++)
            {
                var kaizenCount = monthlyKaizenCounts
                                 .FirstOrDefault(m => m.Month == month)?
                                 .KaizenCount ?? 0;
                dataBar.Add(kaizenCount);
            }
            ViewBag.DataLine = dataBar;

            var monthly5SCounts = await _context.EmployeeViolation5S
                                           .AsNoTracking()
                                           .Where(e => e.DateMonth.Year == currentYear)
                                           .GroupBy(e => e.DateMonth.Month)
                                           .Select(g => new {
                                               Month = g.Key,
                                               V5SCount = g.Count()
                                           })
                                           .ToListAsync();

            var dataLine = new List<int>();
            for (int month = 1; month <= 12; month++)
            {
                var v5sCount = monthly5SCounts
                                 .FirstOrDefault(m => m.Month == month)?
                                 .V5SCount ?? 0;
                dataLine.Add(v5sCount);
            }
            ViewBag.DataBar = dataLine;

            return View();
        }

        [Authorize(Policy = "ViewNotifi")]
        public async Task<IActionResult> GetNotifications()
        {
            var notifications = await _context.Notifications.AsNoTracking().OrderByDescending(n => n.Id).ToListAsync();
            return Json(notifications);
        }

        [Authorize(Policy = "EditNotifi")]
        public async Task<IActionResult> MarkAsRead(long id)
        {
            var notification = await _context.Notifications.FindAsync(id);
            if (notification == null)
            {
                return NotFound();
            }
            notification.IsRead = 1;
            await _context.SaveChangesAsync();
            return Ok();
        }

        [Authorize(Policy = "DeleteNotifi")]
        public async Task<IActionResult> DeleteNotification(long id)
        {
            var notification = await _context.Notifications.FindAsync(id);
            if (notification == null)
            {
                return NotFound();
            }
            _context.Notifications.Remove(notification);
            await _context.SaveChangesAsync();
            return Ok();
        }
    }
}
