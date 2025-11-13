using System.Data;
using ExcelDataReader;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;
using KeyType = (string EmployeeCode, int Year, int Month);

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class ProductivityController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public ProductivityController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewProductivity")]
        public async Task<IActionResult> Index(string key)
        {
            var res = await _context.Productivities.AsNoTracking()
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
        public async Task<IActionResult> Add(Productivity productivity)
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
                return View(productivity);
            }

            var empl = await _context.Employees.AsNoTracking().FirstOrDefaultAsync(e => e.EmployeeCode == productivity.EmployeeCode);
            if (empl == null)
            {
                ModelState.AddModelError("EmployeeCode", "Mã nhân viên không hợp lệ!");
                return View(productivity);
            }

            bool isDuplicate = await IsDuplicateProductivity(
                productivity.EmployeeCode,
                productivity.MeasurementMonth,
                productivity.MeasurementYear
            );
            if (isDuplicate)
            {
                ModelState.AddModelError("EmployeeCode", "Nhân viên đã có năng suất của thời gian này!");
                return View(productivity);
            }

            try
            {
                productivity.EmployeeName = empl.EmployeeName;
                _context.Productivities.Add(productivity);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(productivity);
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

            var productivityToEdit = await _context.Productivities.AsNoTracking().FirstOrDefaultAsync(m => m.Id == id);
            if (productivityToEdit == null)
            {
                return NotFound();
            }
            return View(productivityToEdit);
        }

        [Authorize(Policy = "EditProductivity")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, Productivity productivity)
        {
            if (id != productivity.Id)
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
                return View(productivity);
            }
            var oldProductivity = await _context.Productivities.AsNoTracking().FirstOrDefaultAsync(q => q.Id == productivity.Id);
            if (oldProductivity == null)
            {
                return NotFound();
            }

            var empl = await _context.Employees.AsNoTracking().FirstOrDefaultAsync(e => e.EmployeeCode == productivity.EmployeeCode);
            if (empl == null)
            {
                ModelState.AddModelError("EmployeeCode", "Mã nhân viên không hợp lệ!");
                return View(productivity);
            }
            bool isDuplicate = await IsDuplicateProductivity(
                productivity.EmployeeCode,
                productivity.MeasurementMonth,
                productivity.MeasurementYear,
                productivity.Id
            );
            if (isDuplicate)
            {
                ModelState.AddModelError("EmployeeCode", "Nhân viên đã có năng suất của thời gian này!");
                return View(productivity);
            }
            try
            {
                productivity.EmployeeName = empl.EmployeeName;
                _context.Productivities.Update(productivity);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(productivity);
            }
        }

        [Authorize(Policy = "DeleteProductivity")]
        public async Task<IActionResult> Delete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var productivityToDelete = await _context.Productivities.FindAsync(id);

            if (productivityToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy dữ liệu!";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                _context.Productivities.Remove(productivityToDelete);
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

            var productivityList = new List<Productivity>();
            var requiredFields = new[] { "EmployeeCode", "EmployeeName", "ProductivityScore", "MeasurementYear", "MeasurementMonth" };
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
                            .Select(row => row["EmployeeCode"]?.ToString()?.Trim())
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
                                if (int.TryParse(row["MeasurementYear"]?.ToString(), out int year) &&
                                    int.TryParse(row["MeasurementMonth"]?.ToString(), out int month))
                                {
                                    return new { EmployeeCode = row["EmployeeCode"]?.ToString()?.Trim(), Year = year, Month = month };
                                }
                                return null;
                            })
                            .Where(key => key != null)
                            .Distinct()
                            .ToList();

                        var existingProductivities = await _context.Productivities
                            .Where(p => uniqueKeysInFile.Select(k => k.EmployeeCode).Contains(p.EmployeeCode) &&
                                        uniqueKeysInFile.Select(k => k.Year).Contains(p.MeasurementYear) &&
                                        uniqueKeysInFile.Select(k => k.Month).Contains(p.MeasurementMonth))
                            .Select(p => new KeyType(p.EmployeeCode.Trim(), p.MeasurementYear, p.MeasurementMonth))
                            .ToListAsync();

                        var existingProductivityHashSet = existingProductivities
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
                                var employeeCode = row["EmployeeCode"].ToString().Trim();
                                var year = Convert.ToInt32(row["MeasurementYear"]);
                                var month = Convert.ToInt32(row["MeasurementMonth"]);

                                var currentKey = new KeyType(employeeCode, year, month);

                                if (!existingEmployeeCodes.Contains(employeeCode))
                                {
                                    notFoundEmployeeCount++;
                                    continue;
                                }

                                if (existingProductivityHashSet.Contains(currentKey))
                                {
                                    duplicateCount++;
                                    continue;
                                }

                                var productivity = new Productivity
                                {
                                    EmployeeCode = employeeCode,
                                    EmployeeName = row.Table.Columns.Contains("EmployeeName") && row["EmployeeName"] != DBNull.Value ? row["EmployeeName"].ToString().Trim() : null,
                                    ProductivityScore = Convert.ToDecimal(row["ProductivityScore"]),
                                    MeasurementYear = year,
                                    MeasurementMonth = month
                                };

                                productivityList.Add(productivity);
                            }
                            catch (Exception ex)
                            {
                                return BadRequest(new { Message = $"Lỗi tại dòng {rowIndex}: Lỗi chuyển đổi dữ liệu. Vui lòng kiểm tra định dạng của các cột số (Năng suất, Năm, Tháng). Chi tiết: {ex.Message}" });
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
                if (productivityList.Any())
                {
                    await _context.Productivities.AddRangeAsync(productivityList);
                    await _context.SaveChangesAsync();
                }

                var successCount = productivityList.Count;
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
            var filePath = Path.Combine(_env.WebRootPath, "templates", "import_năng suất.xlsx");
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Không tìm thấy file mẫu.");
            }
            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = "import_năng suất.xlsx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        private async Task<bool> IsDuplicateProductivity(
            string employeeCode,
            int month,
            int year,
            long currentId = 0)
        {
            bool exists = await _context.Productivities
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
