using ExcelDataReader;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using System.Data;
using Web_QM.Models;
using KeyType = (string EmployeeCode, int Year, int Month);

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class SawingPerformanceController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public SawingPerformanceController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewProductivity")]
        public async Task<IActionResult> Index(string key)
        {
            var res = await _context.SawingPerformances.AsNoTracking()
                                            .Where(k =>
                                                string.IsNullOrEmpty(key) ||
                                                k.EmployeeCode.ToLower().Contains(key.ToLower()) ||
                                                k.EmployeeName.ToLower().Contains(key.ToLower())
                                            )
                                            .OrderByDescending(o => o.Id)
                                            .Take(1000)
                                            .ToListAsync();
            ViewBag.KeySearch = key;
            return View(res);
        }

        [Authorize(Policy = "AddProductivity")]
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

        [Authorize(Policy = "AddProductivity")]
        [HttpPost]
        public async Task<IActionResult> Add(SawingPerformance sawingPerformance)
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
                return View(sawingPerformance);
            }

            var empl = await _context.Employees.AsNoTracking().FirstOrDefaultAsync(e => e.EmployeeCode == sawingPerformance.EmployeeCode);
            if (empl == null)
            {
                ModelState.AddModelError(string.Empty, "Mã nhân viên không hợp lệ");
                return View(sawingPerformance);
            }

            bool isDuplicate = await IsDuplicateSawingPerformance(
                sawingPerformance.EmployeeCode,
                sawingPerformance.MeasurementMonth,
                sawingPerformance.MeasurementYear
            );
            if (isDuplicate)
            {
                ModelState.AddModelError(string.Empty, "Nhân viên đã có dữ liệu của thời gian này");
                return View(sawingPerformance);
            }
            decimal rawRate = Math.Round(sawingPerformance.SalesAmountUSD / sawingPerformance.WorkMinute, 2);
            if(rawRate > 999.99m)
            {
                TempData["ErrorMessage"] = "Lỗi, dữ liệu Doanh số/TG làm việc vượt mức cho phép!";
                return View(sawingPerformance);
            }
            sawingPerformance.SalesRate = rawRate;
            try
            {
                sawingPerformance.EmployeeName = empl.EmployeeName;
                _context.SawingPerformances.Add(sawingPerformance);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(sawingPerformance);
            }
        }

        [Authorize(Policy = "EditProductivity")]
        public async Task<IActionResult> Edit(long id)
        {
            if (id == null)
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

            var sawingPerformanceToEdit = await _context.SawingPerformances.AsNoTracking().FirstOrDefaultAsync(m => m.Id == id);
            if (sawingPerformanceToEdit == null)
            {
                return NotFound();
            }
            return View(sawingPerformanceToEdit);
        }

        [Authorize(Policy = "EditProductivity")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, SawingPerformance sawingPerformance)
        {
            if (id != sawingPerformance.Id)
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
                return View(sawingPerformance);
            }
            var oldSawingPerformance = await _context.SawingPerformances.AsNoTracking().FirstOrDefaultAsync(q => q.Id == sawingPerformance.Id);
            if (oldSawingPerformance == null)
            {
                return NotFound();
            }

            var empl = await _context.Employees.AsNoTracking().FirstOrDefaultAsync(e => e.EmployeeCode == sawingPerformance.EmployeeCode);
            if (empl == null)
            {
                ModelState.AddModelError(string.Empty, "Mã nhân viên không hợp lệ");
                return View(sawingPerformance);
            }
            bool isDuplicate = await IsDuplicateSawingPerformance(
                sawingPerformance.EmployeeCode,
                sawingPerformance.MeasurementMonth,
                sawingPerformance.MeasurementYear,
                sawingPerformance.Id
            );
            if (isDuplicate)
            {
                ModelState.AddModelError(string.Empty, "Nhân viên đã có dữ liệu của thời gian này");
                return View(sawingPerformance);
            }
            decimal rawRate = Math.Round(sawingPerformance.SalesAmountUSD / sawingPerformance.WorkMinute, 2);
            if (rawRate > 999.99m)
            {
                TempData["ErrorMessage"] = "Lỗi, dữ liệu Doanh số/TG làm việc vượt mức cho phép!";
                return View(sawingPerformance);
            }
            sawingPerformance.SalesRate = rawRate;
            try
            {
                sawingPerformance.EmployeeName = empl.EmployeeName;
                _context.SawingPerformances.Update(sawingPerformance);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(sawingPerformance);
            }
        }

        [Authorize(Policy = "DeleteProductivity")]
        public async Task<IActionResult> Delete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var sawingPerformanceToDelete = await _context.SawingPerformances.FindAsync(id);

            if (sawingPerformanceToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy dữ liệu!";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                _context.SawingPerformances.Remove(sawingPerformanceToDelete);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa dữ liệu năng suất!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Index));
        }

        [Authorize(Policy = "AddProductivity")]
        [HttpPost]
        public async Task<IActionResult> Import(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest(new { Message = "Vui lòng chọn một file Excel." });
            }

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var sawingPerformanceList = new List<SawingPerformance>();
            var requiredFields = new[] { "Mã NV", "Tên NV", "Doanh số USD", "TG làm việc", "Doanh số/TG", "Năm", "Tháng" };
            var duplicateCount = 0;
            var notFoundEmployeeCount = 0;
            var totalRowsProcessed = 0;

            try
            {
                using (var stream = excelFile.OpenReadStream())
                {
                    using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                    {
                        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                        {
                            return BadRequest(new { Message = "File Excel không có dữ liệu." });
                        }

                        DataTable table = dataSet.Tables[0];
                        totalRowsProcessed = table.Rows.Count;

                        var missingColumns = requiredFields.Where(field => !table.Columns.Contains(field)).ToList();
                        if (missingColumns.Any())
                        {
                            return BadRequest(new { Message = $"File Excel thiếu các cột bắt buộc: {string.Join(", ", missingColumns)}" });
                        }

                        var employeeCodesInFile = table.AsEnumerable()
                            .Select(row => row["Mã NV"]?.ToString()?.Trim())
                            .Where(code => !string.IsNullOrWhiteSpace(code))
                            .Distinct()
                            .ToList();

                        var existingEmployeeCodes = await _context.Employees
                            .Where(e => employeeCodesInFile.Contains(e.EmployeeCode))
                            .Select(e => e.EmployeeCode)
                            .ToListAsync();

                        var uniqueKeysInFile = table.AsEnumerable()
                            .Select(row =>
                            {
                                if (int.TryParse(row["Năm"]?.ToString(), out int year) &&
                                    int.TryParse(row["Tháng"]?.ToString(), out int month))
                                {
                                    return new { EmployeeCode = row["Mã NV"]?.ToString()?.Trim(), Year = year, Month = month };
                                }
                                return null;
                            })
                            .Where(key => key != null)
                            .Distinct()
                            .ToList();

                        var existingSawingPerformances = await _context.SawingPerformances
                            .Where(p => uniqueKeysInFile.Select(k => k.EmployeeCode).Contains(p.EmployeeCode) &&
                                        uniqueKeysInFile.Select(k => k.Year).Contains(p.MeasurementYear) &&
                                        uniqueKeysInFile.Select(k => k.Month).Contains(p.MeasurementMonth))
                            .Select(p => new KeyType(p.EmployeeCode.Trim(), p.MeasurementYear, p.MeasurementMonth))
                            .ToListAsync();

                        var existingSawingPerformanceHashSet = existingSawingPerformances
                            .ToHashSet();

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            var row = table.Rows[i];
                            var rowIndex = i + 2;

                            var missingDataFields = new List<string>();

                            foreach (var field in requiredFields)
                            {
                                if (row[field] == DBNull.Value || string.IsNullOrWhiteSpace(row[field].ToString()))
                                {
                                    missingDataFields.Add(field);
                                }
                            }

                            if (missingDataFields.Any())
                            {
                                return BadRequest(new { Message = $"Lỗi tại dòng {rowIndex}: Các trường bắt buộc bị thiếu dữ liệu: {string.Join(", ", missingDataFields)}." });
                            }

                            try
                            {
                                var employeeCode = row["Mã NV"].ToString().Trim();
                                var year = Convert.ToInt32(row["Năm"]);
                                var month = Convert.ToInt32(row["Tháng"]);

                                var currentKey = new KeyType(employeeCode, year, month);

                                if (!existingEmployeeCodes.Contains(employeeCode))
                                {
                                    notFoundEmployeeCount++;
                                    continue;
                                }

                                if (existingSawingPerformanceHashSet.Contains(currentKey))
                                {
                                    duplicateCount++;
                                    continue;
                                }

                                var sawingPerformance = new SawingPerformance
                                {
                                    EmployeeCode = employeeCode,
                                    EmployeeName = row.Table.Columns.Contains("Tên NV") && row["Tên NV"] != DBNull.Value ? row["Tên NV"].ToString().Trim() : null,
                                    SalesAmountUSD = Convert.ToDecimal(row["Doanh số USD"]),
                                    WorkMinute = Convert.ToInt32(row["TG làm việc"]),
                                    SalesRate = Convert.ToDecimal(row["Doanh số/TG"]),
                                    MeasurementYear = year,
                                    MeasurementMonth = month
                                };

                                sawingPerformanceList.Add(sawingPerformance);
                            }
                            catch (Exception ex)
                            {
                                return BadRequest(new { Message = $"Lỗi tại dòng {rowIndex}: Lỗi chuyển đổi dữ liệu. Vui lòng kiểm tra định dạng của các cột số (Doanh số, TG làm việc, Năm, Tháng). Chi tiết: {ex.Message}" });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { Message = $"Lỗi hệ thống khi đọc file: {ex.Message}" });
            }

            try
            {
                if (sawingPerformanceList.Any())
                {
                    await _context.SawingPerformances.AddRangeAsync(sawingPerformanceList);
                    await _context.SaveChangesAsync();
                }

                var successCount = sawingPerformanceList.Count;
                var message = $"Hoàn tất import. Tổng số dòng trong file: {totalRowsProcessed}. " +
                              $"Đã thêm mới thành công: {successCount} bản ghi. " +
                              $"Đã bỏ qua do trùng lặp: {duplicateCount} bản ghi. " +
                              $"Đã bỏ qua do Mã nhân viên không tồn tại: {notFoundEmployeeCount} bản ghi.";

                return Ok(new { Message = message, SuccessCount = successCount, DuplicateCount = duplicateCount, NotFoundEmployeeCount = notFoundEmployeeCount });
            }
            catch (DbUpdateException ex)
            {
                return StatusCode(500, new { Message = $"Lỗi khi lưu vào Database: Lỗi ràng buộc dữ liệu. Chi tiết: {ex.InnerException?.Message ?? ex.Message}" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { Message = $"Lỗi hệ thống không xác định khi lưu DB: {ex.Message}" });
            }
        }

        [Authorize(Policy = "AddProductivity")]
        public async Task<IActionResult> DownloadTemplate()
        {
            var filePath = Path.Combine(_env.WebRootPath, "templates", "import_doanh số cưa.xlsx");
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Không tìm thấy file mẫu.");
            }
            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = "import_doanh số cưa.xlsx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        private async Task<bool> IsDuplicateSawingPerformance(
            string employeeCode,
            int month,
            int year,
            long currentId = 0)
        {
            bool exists = await _context.SawingPerformances
                .AnyAsync(p =>
                    p.EmployeeCode == employeeCode &&
                    p.MeasurementMonth == month &&
                    p.MeasurementYear == year &&
                    p.Id != currentId
                );
            return exists;
        }
    }
}
