using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;

namespace Web_QM.Controllers
{
    public class SawingPerformanceController : Controller
    {
        private readonly QMContext _context;

        public SawingPerformanceController(QMContext context)
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
            ViewBag.DefaultYear = defaultYear;
            ViewBag.DefaultMonth = defaultMonth;

            return View();
        }

        [Authorize(Policy = "ClientViewProductivity")]
        public async Task<IActionResult> GetPerformanceData(string key, int? month, int? year)
        {
            var query = _context.SawingPerformances
                .Join(
                    _context.Employees,
                    p => p.EmployeeCode,
                    e => e.EmployeeCode,
                    (p, e) => new
                    {
                        Id = p.Id,
                        EmployeeCode = p.EmployeeCode,
                        EmployeeName = e.EmployeeName,
                        Department = e.Department,
                        SalesAmountUSD = p.SalesAmountUSD,
                        WorkMinute = p.WorkMinute,
                        SalesRate = p.SalesRate,
                        MeasurementYear = p.MeasurementYear,
                        MeasurementMonth = p.MeasurementMonth
                    }
                ).AsQueryable();

            if (!string.IsNullOrEmpty(key))
            {
                string searchKey = key.ToLower();
                query = query.Where(p =>
                    p.EmployeeCode.ToLower().Contains(searchKey) ||
                    (p.EmployeeName != null && p.EmployeeName.ToLower().Contains(searchKey))
                );
            }

            if (month.HasValue)
            {
                query = query.Where(p => p.MeasurementMonth == month.Value);
            }

            if (year.HasValue)
            {
                query = query.Where(p => p.MeasurementYear == year.Value);
            }
            var result = await query.OrderByDescending(o => o.SalesRate).ToListAsync();

            return Json(new { success = true, data = result });
        }

        [Authorize(Policy = "ClientViewProductivity")]
        public async Task<IActionResult> GetMonthlyCombinedPerformance(int year)
        {
            if (year < 2010 || year > 2100)
            {
                return BadRequest(new { success = false, message = "Năm không hợp lệ!" });
            }

            try
            {
                var monthlyData = await _context.SawingPerformances
                    .Where(p => p.MeasurementYear == year)
                    .GroupBy(p => p.MeasurementMonth)
                    .Select(g => new
                    {
                        Month = g.Key,
                        AverageSales = g.Average(p => p.SalesRate) ?? 0M,
                        MaxSaleRate = g.Max(p => p.SalesRate) ?? 0M
                    })
                    .OrderBy(m => m.Month)
                    .ToListAsync();

                var fullYearData = Enumerable.Range(1, 12)
                    .GroupJoin(
                        monthlyData,
                        month => month,
                        data => data.Month,
                        (month, dataGroup) => dataGroup.FirstOrDefault() ?? new
                        {
                            Month = month,
                            AverageSales = 0M,
                            MaxSaleRate = 0M
                        }
                    )
                    .OrderBy(d => d.Month)
                    .ToList();

                return Ok(new { success = true, data = fullYearData });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = "Lỗi server khi tải dữ liệu!" });
            }
        }
    }
}
