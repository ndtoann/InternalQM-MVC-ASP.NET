using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Data;
using Web_QM.Models;

namespace Web_QM.Areas.Warehouse.Controllers
{
    [Area("Warehouse")]
    [Authorize]
    public class ToolController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public ToolController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewTool")]
        public async Task<IActionResult> Index()
        {
            return View();
        }

        [Authorize(Policy = "ViewTool")]
        public async Task<IActionResult> GetTools(string type, string key_search)
        {
            var query = _context.Tools.AsQueryable();

            if (!string.IsNullOrEmpty(type))
            {
                query = query.Where(x => x.Type == type);
            }

            if (!string.IsNullOrEmpty(key_search))
            {
                string search = key_search.ToLower();
                query = query.Where(x => x.ToolName.ToLower().Contains(search) || x.ToolCode.ToLower().Contains(search));
            }

            var tools = await query.Take(500).ToListAsync();
            return Json(tools);
        }

        [Authorize(Policy = "AddTool")]
        public async Task<IActionResult> Add()
        {
            return View();
        }

        [Authorize(Policy = "AddTool")]
        [HttpPost]
        public async Task<IActionResult> Add(Tool tool)
        {
            if (!ModelState.IsValid)
            {
                return View(tool);
            }
            var isDuplicate = IsDuplicateToolCode(tool.ToolCode);
            if (isDuplicate)
            {
                ModelState.AddModelError("ToolCode", "Mã vật tư đã tồn tại.");
                return View(tool);
            }
            try
            {
                tool.AvailableQty = tool.InitialQty + tool.TotalImported - tool.TotalScrapped - tool.TotalIssued + tool.TotalReturned;
                tool.CreatedDate = DateOnly.FromDateTime(DateTime.Now);
                _context.Tools.Add(tool);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(tool);
            }
        }

        [Authorize(Policy = "EditTool")]
        public async Task<IActionResult> Edit(long id = 0)
        {
            if (id == 0)
            {
                return NotFound();
            }
            var tool = await _context.Tools.FirstOrDefaultAsync(t => t.Id == id);
            if (tool == null)
            {
                return NotFound();
            }
            return View(tool);
        }

        [Authorize(Policy = "EditTool")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, Tool tool)
        {
            if (id != tool.Id)
            {
                return NotFound();
            }
            if (!ModelState.IsValid)
            {
                return View(tool);
            }
            var isDuplicate = IsDuplicateToolCode(tool.ToolCode, tool.Id);
            if (isDuplicate)
            {
                ModelState.AddModelError("ToolCode", "Mã vật tư đã tồn tại.");
                return View(tool);
            }
            var toolToUpdate = await _context.Tools.AsNoTracking().FirstOrDefaultAsync(t => t.Id == id);
            if (toolToUpdate == null)
            {
                return NotFound();
            }
            try
            {
                tool.AvailableQty = tool.InitialQty + tool.TotalImported - tool.TotalScrapped - tool.TotalIssued + tool.TotalReturned;
                tool.CreatedDate = toolToUpdate.CreatedDate;
                tool.UpdatedDate = DateOnly.FromDateTime(DateTime.Now);
                _context.Tools.Update(tool);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(tool);
            }
        }

        [Authorize(Policy = "DeleteTool")]
        public async Task<IActionResult> Delete(long id = 0)
        {
            if (id == 0)
            {
                return NotFound();
            }
            var tool = await _context.Tools.FindAsync(id);
            if (tool == null)
            {
                return NotFound();
            }
            try
            {
                await _context.ToolSupplyLogs
                     .Where(a => a.ToolId == id)
                     .ExecuteDeleteAsync();

                await _context.IssueReturnLogs
                     .Where(a => a.ToolId == id)
                     .ExecuteDeleteAsync();

                _context.Tools.Remove(tool);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được xóa thành công!";
                return Json(new { success = true, message = "Dữ liệu đã được xóa thành công!" });
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Đã xảy ra lỗi khi xóa dữ liệu. Vui lòng thử lại!";
                return Json(new { success = false, message = "Lỗi hệ thống"});
            }
        }

        [Authorize(Policy = "AddToolSupplyLog")]
        [HttpPost]
        public async Task<IActionResult> SaveSupplyLog(ToolSupplyLog log)
        {
            var employeeCodeIsLogin = User.FindFirst("EmployeeCode")?.Value;
            if (employeeCodeIsLogin == null)
            {
                return Json(new { success = false, message = "Vui lòng đăng nhập" });
            }
            try
            {
                var tool = await _context.Tools.FirstOrDefaultAsync(t => t.Id == log.ToolId);
                if (tool == null) return Json(new { success = false, message = "Vật tư không tồn tại" });

                if (log.Type == "Nhập")
                {
                    tool.TotalImported += log.Qty;
                }
                else if (log.Type == "Hủy")
                {
                    tool.TotalScrapped += log.Qty;
                }
                tool.AvailableQty = tool.InitialQty + tool.TotalImported - tool.TotalScrapped - tool.TotalIssued + tool.TotalReturned;
                tool.UpdatedDate = DateOnly.FromDateTime(DateTime.Now);
                log.WarehouseStaff = employeeCodeIsLogin;
                log.CreatedDate = DateOnly.FromDateTime(DateTime.Now);
                _context.ToolSupplyLogs.Add(log);
                await _context.SaveChangesAsync();

                return Json(new { success = true, message = "Cập nhật số lượng thành công" });
            }
            catch (DbUpdateConcurrencyException)
            {
                return Json(new { success = false, message = "Lỗi hệ thống" });
            }
        }

        [Authorize(Policy = "AddTool")]
        [HttpPost]
        public async Task<IActionResult> Import(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest(new { Message = "Vui lòng chọn một file Excel." });
            }

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var toolList = new List<Tool>();
            var requiredFields = new[] { "Mã vật tư", "Tên vật tư" };
            var duplicateCount = 0;
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

                        var existingToolCodes = await _context.Tools.Select(t => t.ToolCode.Trim()).ToListAsync();
                        var toolCodesInFile = new HashSet<string>();

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            var row = table.Rows[i];
                            var rowIndex = i + 2;

                            if (string.IsNullOrWhiteSpace(row["Mã vật tư"]?.ToString()) || string.IsNullOrWhiteSpace(row["Tên vật tư"]?.ToString()))
                            {
                                return BadRequest(new { Message = $"Lỗi tại dòng {rowIndex}: Mã hoặc Tên vật tư không được để trống." });
                            }

                            var toolCode = row["Mã vật tư"].ToString().Trim();

                            if (existingToolCodes.Contains(toolCode) || toolCodesInFile.Contains(toolCode))
                            {
                                duplicateCount++;
                                continue;
                            }

                            try
                            {
                                int initialQty = int.TryParse(row["Số lượng"]?.ToString(), out int iq) ? iq : 0;
                                int totalImported = int.TryParse(row["SL nhập"]?.ToString(), out int ti) ? ti : 0;
                                int totalScrapped = int.TryParse(row["SL hủy"]?.ToString(), out int ts) ? ts : 0;
                                int totalIssued = int.TryParse(row["SL xuất"]?.ToString(), out int tis) ? tis : 0;
                                int totalReturned = int.TryParse(row["SL trả"]?.ToString(), out int tr) ? tr : 0;

                                var tool = new Tool
                                {
                                    ToolCode = toolCode,
                                    ToolName = row["Tên vật tư"].ToString().Trim(),
                                    Type = table.Columns.Contains("Loại") ? row["Loại"]?.ToString() : null,
                                    Unit = table.Columns.Contains("Đơn vị") ? row["Đơn vị"]?.ToString() : null,
                                    Location = table.Columns.Contains("Vị trí") ? row["Vị trí"]?.ToString() : null,
                                    InitialQty = initialQty,
                                    TotalImported = totalImported,
                                    TotalScrapped = totalScrapped,
                                    TotalIssued = totalIssued,
                                    TotalReturned = totalReturned,
                                    AvailableQty = (initialQty + totalImported - totalScrapped) - totalIssued + totalReturned,
                                    Note = table.Columns.Contains("Ghi chú") ? row["Ghi chú"]?.ToString() : null,
                                    CreatedDate = DateOnly.FromDateTime(DateTime.Now)
                                };

                                toolList.Add(tool);
                                toolCodesInFile.Add(toolCode);
                            }
                            catch (Exception ex)
                            {
                                return BadRequest(new { Message = $"Lỗi tại dòng {rowIndex}: Định dạng số không hợp lệ. Chi tiết: {ex.Message}" });
                            }
                        }
                    }
                }

                if (toolList.Any())
                {
                    await _context.Tools.AddRangeAsync(toolList);
                    await _context.SaveChangesAsync();
                }

                return Ok(new
                {
                    Message = $"Import thành công {toolList.Count} vật tư. Bỏ qua {duplicateCount} mã trùng.",
                    TotalProcessed = totalRowsProcessed,
                    SuccessCount = toolList.Count,
                    DuplicateCount = duplicateCount
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { Message = $"Lỗi hệ thống: {ex.Message}" });
            }
        }

        [Authorize(Policy = "AddTool")]
        public async Task<IActionResult> DownloadTemplate()
        {
            var filePath = Path.Combine(_env.WebRootPath, "templates", "import_vật tư.xlsx");
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Không tìm thấy file mẫu.");
            }
            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = "import_vật tư.xlsx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        [Authorize(Policy = "ViewTool")]
        public async Task<IActionResult> ExportToExcel()
        {
            var tools = await _context.Tools.OrderBy(t => t.ToolCode).ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Data");

                worksheet.Cell(1, 1).Value = "DANH SÁCH VẬT TƯ VÀ TỒN KHO";
                worksheet.Range("A1:N1").Merge().Style.Font.Bold = true;
                worksheet.Range("A1:N1").Style.Font.FontSize = 16;
                worksheet.Range("A1:N1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var headers = new string[]
                {
                "STT", "Mã vật tư", "Tên vật tư", "Loại", "ĐVT", "Vị trí",
                "Số lượng", "SL nhập", "SL hủy", "SL tồn", "SL xuất", "SL trả", "SL khả dụng", "ID"
                };

                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = worksheet.Cell(2, i + 1);
                    cell.Value = headers[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                int row = 3;
                int stt = 1;
                foreach (var item in tools)
                {
                    worksheet.Cell(row, 1).Value = stt++;
                    worksheet.Cell(row, 2).Value = item.ToolCode;
                    worksheet.Cell(row, 3).Value = item.ToolName;
                    worksheet.Cell(row, 4).Value = item.Type;
                    worksheet.Cell(row, 5).Value = item.Unit;
                    worksheet.Cell(row, 6).Value = item.Location;
                    worksheet.Cell(row, 7).Value = item.InitialQty;
                    worksheet.Cell(row, 8).Value = item.TotalImported;
                    worksheet.Cell(row, 9).Value = item.TotalScrapped;
                    worksheet.Cell(row, 10).Value = item.InitialQty + item.TotalImported - item.TotalScrapped;
                    worksheet.Cell(row, 11).Value = item.TotalIssued;
                    worksheet.Cell(row, 12).Value = item.TotalReturned;
                    worksheet.Cell(row, 13).Value = item.AvailableQty;
                    worksheet.Cell(row, 14).Value = item.Id;

                    if (item.AvailableQty <= 0)
                    {
                        worksheet.Cell(row, 14).Style.Font.FontColor = XLColor.Red;
                        worksheet.Cell(row, 14).Style.Font.Bold = true;
                    }

                    worksheet.Range(row, 1, row, 14).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    row++;
                }

                worksheet.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    string fileName = $"Danh sách đồ gá, dao cụ tồn kho.xlsx";

                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName
                    );
                }
            }
        }

        private bool IsDuplicateToolCode(string toolCode, long toolId = 0)
        {
            var existingTool = _context.Tools
                .FirstOrDefault(t => t.ToolCode == toolCode && t.Id != toolId);
            return existingTool != null;
        }
    }
}
