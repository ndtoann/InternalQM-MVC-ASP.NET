using System.Data;
using ExcelDataReader;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;
public record ViolationKey(string EmployeeCode, long Violation5SId, int Year, int Month, int Day);

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class Violation5SController : Controller
    {
        private readonly QMContext _context;

        public Violation5SController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ViewViolation5S")]
        public async Task<IActionResult> Index(string key)
        {
            var res = await _context.Violation5S.AsNoTracking()
                                            .Where(k =>
                                                string.IsNullOrEmpty(key) ||
                                                k.Content5S.ToLower().Contains(key.ToLower())
                                            )
                                            .ToListAsync();
            ViewBag.KeySearch = key;
            return View(res);
        }

        [Authorize(Policy = "AddViolation5S")]
        public async Task<IActionResult> Add()
        {
            return View();
        }

        [Authorize(Policy = "AddViolation5S")]
        [HttpPost]
        public async Task<IActionResult> Add(Violation5S violation5S)
        {
            if (!ModelState.IsValid)
            {
                return View(violation5S);
            }
            try
            {
                violation5S.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Violation5S.Add(violation5S);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(violation5S);
            }
        }

        [Authorize(Policy = "EditViolation5S")]
        public async Task<IActionResult> Edit(long id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var violation5S = await _context.Violation5S.AsNoTracking().FirstOrDefaultAsync(v => v.Id == id);
            if (violation5S == null)
            {
                return NotFound();
            }
            return View(violation5S);
        }

        [Authorize(Policy = "EditViolation5S")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, Violation5S violation5S)
        {
            if (id != violation5S.Id)
            {
                return NotFound();
            }
            var violation5SToEdit = await _context.Violation5S.AsNoTracking().FirstOrDefaultAsync(v => v.Id == id);
            if (violation5SToEdit == null)
            {
                return NotFound();
            }
            try
            {
                violation5S.UpdatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Violation5S.Update(violation5S);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(violation5S);
            }
        }

        [Authorize(Policy = "DeleteViolation5S")]
        public async Task<IActionResult> Delete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var violation5SToDelete = await _context.Violation5S.FindAsync(id);
            if (violation5SToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy dữ liệu!";
                return RedirectToAction(nameof(Index));
            }
            try
            {
                await _context.EmployeeViolation5S
                     .Where(a => a.Violation5SId == violation5SToDelete.Id)
                     .ExecuteDeleteAsync();

                _context.Violation5S.Remove(violation5SToDelete);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa dữ liệu!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Index));
        }

        [Authorize(Policy = "EmplViolation5S")]
        [HttpPost]
        public async Task<IActionResult> Import(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest(new { Message = "Vui lòng chọn một file Excel." });
            }

            var violationsToSave = new List<EmployeeViolation5S>();
            var requiredFields = new[] { "EmployeeCode", "Violation5SId", "DateMonth", "Qty" };
            var duplicateCount = 0;
            var notFoundViolationCount = 0;
            var totalRowsProcessed = 0;
            var errors = new Dictionary<int, string>();
            var violationsToUpdate = new List<EmployeeViolation5S>();

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

                        var violationIdsInFile = table.AsEnumerable()
                            .Select(row => row["Violation5SId"]?.ToString()?.Trim())
                            .Where(idStr => !string.IsNullOrWhiteSpace(idStr) && long.TryParse(idStr, out _))
                            .Select(idStr => long.Parse(idStr))
                            .Distinct()
                            .ToList();

                        var existingViolationIds = await _context.Violation5S
                            .Where(v => violationIdsInFile.Contains(v.Id))
                            .Select(v => v.Id)
                            .ToListAsync();

                        var uniqueKeysInFile = table.AsEnumerable()
                            .Select(row =>
                            {
                                if (long.TryParse(row["Violation5SId"]?.ToString(), out long vid) &&
                                    DateTime.TryParse(row["DateMonth"]?.ToString(), out DateTime date))
                                {
                                    return new ViolationKey(
                                        row["EmployeeCode"]?.ToString()?.Trim(),
                                        vid,
                                        date.Year,
                                        date.Month,
                                        date.Day
                                    );
                                }
                                return null;
                            })
                            .Where(key => key != null)
                            .Distinct()
                            .ToList();

                        var existingViolationList = await _context.EmployeeViolation5S
                        .Where(p => uniqueKeysInFile.Select(k => k.EmployeeCode).Contains(p.EmployeeCode) &&
                                    uniqueKeysInFile.Select(k => k.Violation5SId).Contains(p.Violation5SId) &&
                                    uniqueKeysInFile.Select(k => k.Year).Contains(p.DateMonth.Year) &&
                                    uniqueKeysInFile.Select(k => k.Month).Contains(p.DateMonth.Month) &&
                                    uniqueKeysInFile.Select(k => k.Day).Contains(p.DateMonth.Day))
                        .Select(p => new { Violation = p, Key = new ViolationKey(p.EmployeeCode.Trim(), p.Violation5SId, p.DateMonth.Year, p.DateMonth.Month, p.DateMonth.Day) })
                        .ToListAsync();

                        var existingViolationDictionary = existingViolationList.ToDictionary(x => x.Key, x => x.Violation);

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
                                errors.Add(rowIndex, $"Các trường bắt buộc bị thiếu dữ liệu: {string.Join(", ", missingDataFields)}.");
                                continue;
                            }

                            try
                            {
                                var employeeCode = row["EmployeeCode"].ToString().Trim();
                                var violationId = Convert.ToInt64(row["Violation5SId"]);
                                var dateMonthStr = row["DateMonth"].ToString().Trim();
                                var qtyStr = row["Qty"].ToString().Trim();

                                DateTime dateMonth;
                                if (double.TryParse(dateMonthStr, out double excelDateValue))
                                {
                                    dateMonth = DateTime.FromOADate(excelDateValue);
                                }
                                else if (!DateTime.TryParse(dateMonthStr, out dateMonth))
                                {
                                    errors.Add(rowIndex, $"Định dạng Ngày/Tháng ('{dateMonthStr}') không hợp lệ. Vui lòng sử dụng định dạng ngày hợp lệ.");
                                    continue;
                                }

                                int qty;
                                if (!int.TryParse(qtyStr, out qty) || qty <= 0)
                                {
                                    errors.Add(rowIndex, $"Số lượng lỗi ('{qtyStr}') không hợp lệ. Phải là số nguyên dương.");
                                    continue;
                                }

                                if (!existingViolationIds.Contains(violationId))
                                {
                                    notFoundViolationCount++;
                                    errors.Add(rowIndex, $"Violation5SId '{violationId}' không tồn tại trong hệ thống.");
                                    continue;
                                }

                                var currentKey = new ViolationKey(employeeCode, violationId, dateMonth.Year, dateMonth.Month, dateMonth.Day);

                                if (existingViolationDictionary.TryGetValue(currentKey, out var existingViolation))
                                {
                                    duplicateCount++;
                                    existingViolation.Qty += qty;
                                    violationsToUpdate.Add(existingViolation);
                                }
                                else
                                {
                                    existingViolationDictionary.Add(currentKey, new EmployeeViolation5S());

                                    var violation = new EmployeeViolation5S
                                    {
                                        EmployeeCode = employeeCode,
                                        Violation5SId = violationId,
                                        DateMonth = DateOnly.FromDateTime(dateMonth),
                                        Qty = qty
                                    };

                                    violationsToSave.Add(violation);
                                }
                            }
                            catch (Exception ex)
                            {
                                errors.Add(rowIndex, $"Lỗi chuyển đổi dữ liệu. Chi tiết: {ex.Message}");
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
                var successAddCount = violationsToSave.Count;
                var successUpdateCount = violationsToUpdate.Distinct().Count();

                if (successAddCount > 0)
                {
                    await _context.EmployeeViolation5S.AddRangeAsync(violationsToSave);
                }

                if (successUpdateCount > 0)
                {
                    _context.EmployeeViolation5S.UpdateRange(violationsToUpdate.Distinct());
                }

                if (successAddCount > 0 || successUpdateCount > 0)
                {
                    await _context.SaveChangesAsync();
                }

                var finalMessage = $"Hoàn tất import. Tổng số dòng trong file: {totalRowsProcessed}. " +
                                    $"Đã thêm mới thành công: {successAddCount} bản ghi. " +
                                    $"Đã cập nhật số lần lỗi cho: {successUpdateCount} bản ghi trùng lặp. " +
                                    $"Đã bỏ qua do Mã Lỗi 5S không tồn tại: {notFoundViolationCount} bản ghi.";

                if (errors.Any())
                {
                    return Ok(new
                    {
                        Message = finalMessage,
                        Warning = $"Có {errors.Count} dòng bị lỗi và đã bị bỏ qua.",
                        SuccessCount = successAddCount,
                        UpdateCount = successUpdateCount,
                        NotFoundViolationCount = notFoundViolationCount,
                        Errors = errors.Select(e => $"Dòng {e.Key}: {e.Value}").ToList()
                    });
                }

                return Ok(new
                {
                    Message = finalMessage,
                    SuccessCount = successAddCount,
                    UpdateCount = successUpdateCount,
                    NotFoundViolationCount = notFoundViolationCount
                });
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
    }
}