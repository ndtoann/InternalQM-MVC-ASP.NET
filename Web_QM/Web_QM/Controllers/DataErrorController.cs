using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using ExcelDataReader;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Data;
using System.Globalization;
using System.Threading.Tasks;
using Web_QM.Models;
using Web_QM.Models.ViewModels;

namespace Web_QM.Controllers
{
    public class DataErrorController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public DataErrorController( QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ClientViewProductionDefect")]
        public async Task<IActionResult> Index()
        {
            return View();
        }

        [Authorize(Policy = "ClientViewProductionDefect")]
        public async Task<IActionResult> GetDataErrors(string key, DateTime? startDate = null, DateTime? endDate = null)
        {
            var res = await _context.ErrorDatas.AsNoTracking()
                                            .Where(k =>
                                                (string.IsNullOrEmpty(key) ||
                                                (k.EmployeeCode.ToLower().Contains(key.ToLower())) ||
                                                (k.OrderNo.ToLower().Contains(key.ToLower())) ||
                                                (k.PartName.ToLower().Contains(key.ToLower()))) &&
                                                (k.DateMonth >= DateOnly.FromDateTime(startDate.Value)) &&
                                                (k.DateMonth <= DateOnly.FromDateTime(endDate.Value))
                                            )
                                            .Select(e => new ErrorDataView
                                            {
                                                Id = e.Id,
                                                OrderNo = e.OrderNo,
                                                PartName = e.PartName,
                                                QtyOrder = e.QtyOrder,
                                                QtyNG = e.QtyNG,
                                                DateMonth = e.DateMonth,
                                                ErrorDetected = e.ErrorDetected,
                                                ErrorType = e.ErrorType,
                                                NCC = e.NCC,
                                                EmployeeCode = e.EmployeeCode,
                                                Department = e.Department,
                                                ErrorCompletionDate = e.ErrorCompletionDate
                                            })
                                            .OrderByDescending(o => o.DateMonth)
                                            .Take(1000)
                                            .ToListAsync();

            var dataResult = res.Select(e => new
            {
                e.Id,
                e.OrderNo,
                e.PartName,
                e.QtyOrder,
                e.QtyNG,
                NGRate = e.QtyOrder > 0 ? (e.QtyNG / (double)e.QtyOrder) * 100 : 0,
                DateMonth = e.DateMonth.ToString("dd/MM/yyyy"),
                e.ErrorDetected,
                e.ErrorType,
                e.NCC,
                e.EmployeeCode,
                e.Department,
                ErrorCompletionDate = e.ErrorCompletionDate.HasValue ? e.ErrorCompletionDate.Value.ToString("dd/MM/yyyy") : null,
            });

            return Json(new { data = dataResult, total = dataResult.Count() });
        }

        [Authorize(Policy = "ClientViewProductionDefect")]
        public async Task<IActionResult> Getdetail(long id)
        {
            var data = await _context.ErrorDatas.AsNoTracking()
                                              .FirstOrDefaultAsync(e => e.Id == id);
            if (data == null)
            {
                return NotFound();
            }
            return Json(data);
        }

        [Authorize(Policy = "ViewProductionDefect")]
        public async Task<IActionResult> GetErrorChartData(int year)
        {
            int currentYear = year;

            var errors = await _context.ErrorDatas
                                       .Where(e => e.DateMonth.Year == currentYear)
                                       .ToListAsync();

            var monthlyData = errors.GroupBy(e => e.DateMonth.Month)
                                    .OrderBy(g => g.Key)
                                    .ToDictionary(g => g.Key, g => g.ToList());

            var chartData = new List<object>();

            for (int month = 1; month <= 12; month++)
            {
                int qmErrorsCount = 0;
                int otherNccErrorsCount = 0;

                if (monthlyData.ContainsKey(month))
                {
                    var monthErrors = monthlyData[month];

                    qmErrorsCount = monthErrors.Count(e => e.NCC.Trim().ToUpper() == "QM");
                    otherNccErrorsCount = monthErrors.Count(e => e.NCC.Trim().ToUpper() != "QM");
                }

                var monthData = new
                {
                    Month = $"T{month}",
                    QMErrors = qmErrorsCount,
                    OtherNccErrors = otherNccErrorsCount
                };
                chartData.Add(monthData);
            }

            return Ok(chartData);
        }

        [Authorize(Policy = "ViewProductionDefect")]
        public async Task<IActionResult> GetLineChartData(int year)
        {
            var currentYear = year;

            var data = await _context.ErrorDatas
                .Where(e => e.DateMonth.Year == currentYear)
                .GroupBy(e => new { Ncc = e.NCC, Month = e.DateMonth.Month })
                .Select(g => new
                {
                    NCC = g.Key.Ncc,
                    Month = g.Key.Month,
                    TotalErrors = g.Count()
                })
                .ToListAsync();

            var nccs = data.Select(d => d.NCC).Distinct().ToList();

            var datasets = nccs.Select(ncc => new
            {
                label = ncc,
                data = Enumerable.Range(1, 12)
                    .Select(month => data.FirstOrDefault(d => d.NCC == ncc && d.Month == month)?.TotalErrors ?? 0)
                    .ToList()
            }).ToList();

            var chartData = new
            {
                labels = Enumerable.Range(1, 12).Select(m => $"Tháng {m}").ToList(),
                datasets = datasets
            };

            return Ok(chartData);
        }

        [Authorize(Policy = "ClientViewViolation5S")]
        public async Task<IActionResult> Violation5S()
        {
            ViewBag.Departments = await _context.Departments
                                                .Select(d => new { d.DepartmentName })
                                                .ToListAsync();
            return View();
        }

        [Authorize(Policy = "ClientViewViolation5S")]
        public async Task<IActionResult> GetViolations(string key, string department, DateTime? startDate = null, DateTime? endDate = null)
        {
            try
            {
                string searchKey = (key ?? "").ToLower();
                string filterDepartment = department ?? "";
                DateOnly? startOnly = startDate.HasValue ? DateOnly.FromDateTime(startDate.Value.Date) : (DateOnly?)null;
                DateOnly? endOnly = endDate.HasValue ? DateOnly.FromDateTime(endDate.Value.Date) : (DateOnly?)null;

                var query = _context.EmployeeViolation5S
                    .AsNoTracking()
                    .Join(_context.Employees,
                          v => v.EmployeeCode,
                          e => e.EmployeeCode,
                          (v, e) => new { Violation = v, Employee = e })
                    .Join(_context.Violation5S,
                          ve => ve.Violation.Violation5SId,
                          vs => vs.Id,
                          (ve, vs) => new { ve.Violation, ve.Employee, ViolationContent = vs });

                query = query.Where(x =>
                    (string.IsNullOrEmpty(searchKey) ||
                     (x.Employee.EmployeeCode != null && x.Employee.EmployeeCode.ToLower().Contains(searchKey)) ||
                     (x.Employee.EmployeeName != null && x.Employee.EmployeeName.ToLower().Contains(searchKey))
                    )
                    &&
                    (string.IsNullOrEmpty(filterDepartment) || x.Employee.Department == filterDepartment)
                    &&
                    (!startOnly.HasValue || x.Violation.DateMonth >= startOnly.Value)
                    &&
                    (!endOnly.HasValue || x.Violation.DateMonth <= endOnly.Value)
                );

                var violations = await query
                    .OrderByDescending(x => x.Violation.DateMonth)
                    .Take(1000)
                    .Select(x => new
                    {
                        Id = x.Violation.Id,
                        EmployeeCode = x.Violation.EmployeeCode,
                        EmployeeName = x.Employee.EmployeeName,
                        Department = x.Employee.Department,
                        Content5S = x.ViolationContent.Content5S,
                        DateMonth = x.Violation.DateMonth,
                        Qty = x.Violation.Qty
                    })
                    .ToListAsync();

                return Json(new { success = true, data = violations });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }

        [Authorize(Policy = "ClientViewViolation5S")]
        public async Task<IActionResult> GetChartData(int year)
        {
            try
            {
                var monthlyData = await _context.EmployeeViolation5S
                    .Where(v => v.DateMonth.Year == year)
                    .GroupBy(v => v.DateMonth.Month)
                    .Select(g => new 
                    {
                        Month = g.Key,
                        TotalQty = g.Sum(v => v.Qty)
                    })
                    .OrderBy(m => m.Month)
                    .ToListAsync();

                var departmentData = await _context.EmployeeViolation5S
                    .Where(v => v.DateMonth.Year == year)
                    .Join(_context.Employees,
                          v => v.EmployeeCode,
                          e => e.EmployeeCode,
                          (v, e) => new { Violation = v, Employee = e })
                    .Where(x => x.Employee.Department != null)
                    .GroupBy(x => x.Employee.Department)
                    .Select(g => new
                    {
                        Department = g.Key,
                        TotalQty = g.Sum(x => x.Violation.Qty)
                    })
                    .OrderByDescending(d => d.TotalQty)
                    .ToListAsync();

                var chartData = new
                {
                    MonthlyData = monthlyData,
                    DepartmentData = departmentData
                };

                return Json(new { success = true, data = chartData });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = $"Lỗi khi tải dữ liệu biểu đồ: {ex.Message}" });
            }
        }
    }
}
