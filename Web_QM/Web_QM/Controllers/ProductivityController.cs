using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;

namespace Web_QM.Controllers
{
    public class ProductivityController : Controller
    {
        private readonly QMContext _context;

        public ProductivityController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ClientViewProductivity")]
        public async Task<IActionResult> Index()
        {
            var today = DateTime.Today;
            int defaultMonth = today.Month - 1;
            int defaultYear = today.Year;
            if (defaultMonth == 0)
            {
                defaultMonth = 12;
                defaultYear = today.Year - 1;
            }
            ViewBag.Departments = await _context.Departments.ToListAsync();
            ViewBag.DefaultYear = defaultYear;
            ViewBag.DefaultMonth = defaultMonth;

            return View();
        }

        [Authorize(Policy = "ClientViewProductivity")]
        public async Task<IActionResult> GetProductivityData(
        string key,
        string department,
        int? year,
        int? month)
        {
            try
            {
                string searchKey = (key ?? "").ToLower();
                string filterDepartment = department ?? "";
                int filterYear = year ?? DateTime.Today.Year;
                int? filterMonth = month;

                var query = _context.Productivities
                    .AsNoTracking()
                    .GroupJoin(_context.Employees,
                        p => p.EmployeeCode,
                        e => e.EmployeeCode,
                        (p, employees) => new { Productivity = p, Employees = employees.DefaultIfEmpty() })
                    .SelectMany(
                        x => x.Employees,
                        (x, e) => new { Productivity = x.Productivity, Employee = e }
                    );

                query = query.Where(x =>
                    (string.IsNullOrEmpty(searchKey) ||
                     (x.Productivity.EmployeeCode != null && x.Productivity.EmployeeCode.ToLower().Contains(searchKey)) ||
                     (x.Employee != null && x.Employee.EmployeeName != null && x.Employee.EmployeeName.ToLower().Contains(searchKey))
                    )
                    &&
                    (string.IsNullOrEmpty(filterDepartment) ||
                     (x.Employee != null && x.Employee.Department == filterDepartment)
                    )
                    &&
                    (x.Productivity.MeasurementYear == filterYear)
                    &&
                    (!filterMonth.HasValue || x.Productivity.MeasurementMonth == filterMonth.Value)
                );

                var results = await query
                    .OrderByDescending(x => x.Productivity.MeasurementYear)
                    .ThenByDescending(x => x.Productivity.MeasurementMonth)
                    .Select(x => new
                    {
                        Id = x.Productivity.Id,
                        EmployeeCode = x.Productivity.EmployeeCode,
                        EmployeeName = x.Employee != null ? x.Employee.EmployeeName : null,
                        Department = x.Employee != null ? x.Employee.Department : null,
                        ProductivityScore = x.Productivity.ProductivityScore,
                        MeasurementYear = x.Productivity.MeasurementYear,
                        MeasurementMonth = x.Productivity.MeasurementMonth
                    })
                    .ToListAsync();

                return Json(new { success = true, data = results });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }

        [Authorize(Policy = "ClientViewProductivity")]
        public async Task<IActionResult> GetMonthlyAverageProductivity(int year)
        {
            try
            {
                var data = await _context.Productivities
                    .Where(p => p.MeasurementYear == year)
                    .GroupBy(p => p.MeasurementMonth)
                    .Select(g => new
                    {
                        Month = g.Key,
                        AverageScore = g.Average(p => (double)p.ProductivityScore)
                    })
                    .OrderBy(g => g.Month)
                    .ToListAsync();

                var results = Enumerable.Range(1, 12).Select(month => {
                    var monthData = data.FirstOrDefault(d => d.Month == month);
                    double score = monthData?.AverageScore ?? 0;
                    double cappedScore = Math.Min(score, 200.0);
                    return new
                    {
                        Month = month,
                        AverageScore = cappedScore
                    };
                }).ToList();

                return Json(new { success = true, data = results });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }

        [Authorize(Policy = "ClientViewProductivity")]
        public async Task<IActionResult> GetDepartmentProductivityByMonth(int year, int month)
        {
            try
            {
                var data = await _context.Productivities
                    .Where(p => p.MeasurementYear == year && p.MeasurementMonth == month)
                    .Join(_context.Employees,
                          p => p.EmployeeCode,
                          e => e.EmployeeCode,
                          (p, e) => new { Productivity = p, Employee = e })
                    .Where(x => x.Employee.Department != null)
                    .GroupBy(x => x.Employee.Department)
                    .Select(g => new
                    {
                        Department = g.Key,
                        AverageScore = g.Average(x => (double)x.Productivity.ProductivityScore)
                    })
                    .OrderBy(g => g.Department)
                    .ToListAsync();

                var results = data.Select(d => new
                {
                    Department = d.Department,
                    AverageScore = Math.Min(d.AverageScore, 200.0)
                }).ToList();

                return Json(new { success = true, data = results });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }
    }
}
