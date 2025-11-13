using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDataReader;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using System.Data;
using Web_QM.Models;

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class MachineMaintenanceController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public MachineMaintenanceController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewMachineMaintenance")]
        public async Task<IActionResult> Index(string key, int isComplete = -1, DateTime? startDate = null, DateTime? endDate = null)
        {
            if (startDate == null)
            {
                startDate = DateTime.Today.AddYears(-5);
            }
            if (endDate == null)
            {
                endDate = DateTime.Today;
            }
            endDate = endDate.Value.Date.AddDays(1).AddSeconds(-1);

            var res = await _context.MachineMaintenances.AsNoTracking()
                .Where(m =>
                    (isComplete == -1 || m.IsComplete == isComplete) &&
                    (string.IsNullOrEmpty(key) || m.MachineCode.ToLower().Contains(key.ToLower())) &&
                    (m.DateMonth >= DateOnly.FromDateTime(startDate.Value)) &&
                    (m.DateMonth <= DateOnly.FromDateTime(endDate.Value))
                )
                .OrderByDescending(o => o.DateMonth)
                .Take(1000)
                .ToListAsync();

            var statusOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "-1", Text = "Tất cả" },
                new SelectListItem { Value = "1", Text = "Hoàn thành" },
                new SelectListItem { Value = "0", Text = "Chưa hoàn thành" }
            };

            ViewData["IsCompleteList"] = new SelectList(statusOptions, "Value", "Text", isComplete);
            ViewBag.KeySearch = key;
            ViewBag.StartDate = startDate.Value.ToString("yyyy-MM-dd");
            ViewBag.EndDate = endDate.Value.Date.ToString("yyyy-MM-dd");

            return View(res);
        }

        [Authorize(Policy = "AddMachineMaintenance")]
        public async Task<IActionResult> Add()
        {
            var machines = await _context.Machines.ToListAsync();
            var machineList = machines.Select(d => new SelectListItem
            {
                Value = d.MachineCode,
                Text = d.MachineCode + " - " + d.MachineName
            }).ToList();
            machineList.Insert(0, new SelectListItem { Value = "", Text = "Chọn máy" });
            ViewData["Machines"] = machineList;
            return View();
        }

        [Authorize(Policy = "AddMachineMaintenance")]
        [HttpPost]
        public async Task<IActionResult> Add(MachineMaintenance machineMaintenance)
        {
            var machines = await _context.Machines.ToListAsync();
            var machineList = machines.Select(d => new SelectListItem
            {
                Value = d.MachineCode,
                Text = d.MachineCode + " - " + d.MachineName
            }).ToList();
            machineList.Insert(0, new SelectListItem { Value = "", Text = "Chọn máy" });
            ViewData["Machines"] = machineList;
            if (!ModelState.IsValid)
            {
                return View(machineMaintenance);
            }
            try
            {
                _context.MachineMaintenances.Add(machineMaintenance);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!";
                return View(machineMaintenance);
            }
        }

        [Authorize(Policy = "EditMachineMaintenance")]
        public async Task<IActionResult> Edit(long id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var mmToEdit = await _context.MachineMaintenances.FindAsync(id);
            if (mmToEdit == null)
            {
                return NotFound();
            }
            var machines = await _context.Machines.ToListAsync();
            var machineList = machines.Select(d => new SelectListItem
            {
                Value = d.MachineCode,
                Text = d.MachineCode + " - " + d.MachineName
            }).ToList();
            machineList.Insert(0, new SelectListItem { Value = "", Text = "Chọn máy" });
            ViewData["Machines"] = machineList;
            return View(mmToEdit);
        }

        [Authorize(Policy = "EditMachineMaintenance")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, MachineMaintenance machineMaintenance)
        {
            if (id != machineMaintenance.Id)
            {
                return NotFound();
            }
            var machines = await _context.Machines.ToListAsync();
            var machineList = machines.Select(d => new SelectListItem
            {
                Value = d.MachineCode,
                Text = d.MachineCode + " - " + d.MachineName
            }).ToList();
            machineList.Insert(0, new SelectListItem { Value = "", Text = "Chọn máy" });
            ViewData["Machines"] = machineList;
            if (!ModelState.IsValid)
            {
                return View(machineMaintenance);
            }
            var mmToEdit = await _context.MachineMaintenances.AsNoTracking().FirstOrDefaultAsync(x => x.Id == machineMaintenance.Id);
            if (mmToEdit == null)
            {
                return NotFound();
            }
            try
            {
                machineMaintenance.IsComplete = mmToEdit.IsComplete;
                _context.MachineMaintenances.Update(machineMaintenance);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!";
                return View(machineMaintenance);
            }
        }

        [Authorize(Policy = "DeleteMachineMaintenance")]
        public async Task<IActionResult> Delete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var mmToDelete = await _context.MachineMaintenances.FirstOrDefaultAsync(x => x.Id == id);
            if (mmToDelete == null)
            {
                return NotFound();
            }
            try
            {
                _context.MachineMaintenances.Remove(mmToDelete);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa thành công!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }
            return RedirectToAction(nameof(Index));
        }

        [Authorize(Policy = "EditMachineMaintenance")]
        public async Task<IActionResult> ConfirmComplete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var maintenance = await _context.MachineMaintenances.FirstOrDefaultAsync(m => m.Id == id);
            if (maintenance == null)
            {
                return NotFound();
            }
            try
            {
                maintenance.IsComplete = 1;
                _context.MachineMaintenances.Update(maintenance);
                await _context.SaveChangesAsync();
                return Json(new { success = true, message = "Cập nhật thành công!" });
            }
            catch (DbUpdateConcurrencyException)
            {
                return Json(new { success = false, message = "Xảy ra lỗi, vui lòng thử lại!" });
            }
        }

        [Authorize(Policy = "AddMachineMaintenance")]
        [HttpPost]
        public async Task<IActionResult> Import(IFormFile excelFile)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest(new { Message = "Vui lòng chọn một file Excel." });
            }

            var maintenancesToSave = new List<MachineMaintenance>();
            var requiredFields = new[] { "MachineCode", "DateMonth", "MaintenanceContent" };
            var duplicateCount = 0;
            var notFoundMachineCount = 0;
            var totalRowsProcessed = 0;
            var errors = new Dictionary<int, string>();

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

                        var machineCodesInFile = table.AsEnumerable()
                            .Select(row => row["MachineCode"]?.ToString()?.Trim())
                            .Where(code => !string.IsNullOrWhiteSpace(code))
                            .Distinct()
                            .ToList();

                        var existingMachineCodes = await _context.Machines
                            .Where(m => machineCodesInFile.Contains(m.MachineCode))
                            .Select(m => m.MachineCode)
                            .ToListAsync();

                        var existingMaintenanceKeys = await _context.MachineMaintenances
                            .Select(m => new { m.MachineCode, m.MaintenanceContent, m.DateMonth })
                            .ToListAsync();

                        var existingMaintenanceDictionary = existingMaintenanceKeys
                            .ToDictionary(k => (k.MachineCode, k.MaintenanceContent, k.DateMonth), k => true);


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
                                var machineCode = row["MachineCode"].ToString().Trim();
                                var maintenanceContent = row["MaintenanceContent"].ToString().Trim();
                                var dateMonthStr = row["DateMonth"].ToString().Trim();

                                var maintenanceStaff = row.Table.Columns.Contains("MaintenanceStaff") ? row["MaintenanceStaff"]?.ToString()?.Trim() : null;
                                var note = row.Table.Columns.Contains("Note") ? row["Note"]?.ToString()?.Trim() : null;
                                var isCompleteStr = row["IsComplete"].ToString().Trim();

                                DateTime tempDate;
                                if (double.TryParse(dateMonthStr, out double excelDateValue))
                                {
                                    tempDate = DateTime.FromOADate(excelDateValue);
                                }
                                else if (!DateTime.TryParse(dateMonthStr, out tempDate))
                                {
                                    errors.Add(rowIndex, $"Định dạng Ngày/Tháng ('{dateMonthStr}') không hợp lệ. Vui lòng sử dụng định dạng ngày hợp lệ.");
                                    continue;
                                }

                                var dateMonth = DateOnly.FromDateTime(tempDate);

                                if (!existingMachineCodes.Contains(machineCode))
                                {
                                    notFoundMachineCount++;
                                    errors.Add(rowIndex, $"Mã máy '{machineCode}' không tồn tại trong hệ thống.");
                                    continue;
                                }

                                var currentKey = (MachineCode: machineCode, MaintenanceContent: maintenanceContent, DateMonth: dateMonth);

                                if (existingMaintenanceDictionary.ContainsKey(currentKey))
                                {
                                    duplicateCount++;
                                    errors.Add(rowIndex, $"Bản ghi bị trùng lặp trong hệ thống.");
                                    continue;
                                }
                                if (maintenancesToSave.Any(m => m.MachineCode == machineCode && m.MaintenanceContent == maintenanceContent && m.DateMonth == dateMonth))
                                {
                                    duplicateCount++;
                                    errors.Add(rowIndex, $"Bản ghi bị trùng lặp trong file.");
                                    continue;
                                }

                                int isComplete;
                                if (!int.TryParse(isCompleteStr, out isComplete) || isComplete < 0)
                                {
                                    errors.Add(rowIndex, $"Thuộc tính isComplete: ('{isCompleteStr}') không hợp lệ!");
                                    continue;
                                }

                                var maintenance = new MachineMaintenance
                                {
                                    MachineCode = machineCode,
                                    DateMonth = dateMonth,
                                    MaintenanceContent = maintenanceContent,
                                    MaintenanceStaff = maintenanceStaff,
                                    Note = note,
                                    IsComplete = isComplete
                                };

                                maintenancesToSave.Add(maintenance);
                                existingMaintenanceDictionary.Add(currentKey, true);
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
                var successAddCount = maintenancesToSave.Count;

                if (successAddCount > 0)
                {
                    await _context.MachineMaintenances.AddRangeAsync(maintenancesToSave);
                    await _context.SaveChangesAsync();
                }

                var finalMessage = $"Hoàn tất import. Tổng số dòng trong file: {totalRowsProcessed}. " +
                                    $"Đã thêm mới thành công: {successAddCount} bản ghi. " +
                                    $"Đã bỏ qua do bị trùng lặp: {duplicateCount} bản ghi. " +
                                    $"Đã bỏ qua do Mã máy không tồn tại: {notFoundMachineCount} bản ghi.";

                if (errors.Any())
                {
                    return Ok(new
                    {
                        Message = finalMessage,
                        SuccessCount = successAddCount,
                        NotFoundMachineCount = notFoundMachineCount,
                        DuplicateCount = duplicateCount,
                        Errors = errors.Select(e => $"Dòng {e.Key}: {e.Value}").ToList()
                    });
                }

                return Ok(new
                {
                    Message = finalMessage,
                    SuccessCount = successAddCount,
                    NotFoundMachineCount = notFoundMachineCount,
                    DuplicateCount = duplicateCount
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

        [Authorize(Policy = "AddMachineMaintenance")]
        public async Task<IActionResult> DownloadTemplate()
        {
            var filePath = Path.Combine(_env.WebRootPath, "templates", "import_kế hoạch bảo dưỡng máy.xlsx");
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Không tìm thấy file mẫu.");
            }
            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = "import_kế hoạch bảo dưỡng máy.xlsx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }
}
