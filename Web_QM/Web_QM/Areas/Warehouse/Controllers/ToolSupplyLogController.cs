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
    public class ToolSupplyLogController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public ToolSupplyLogController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewToolSupplyLog")]
        public async Task<IActionResult> Index()
        {
            return View();
        }

        [Authorize(Policy = "ViewToolSupplyLog")]
        public async Task<IActionResult> GetToolSupplyLogs(string? type, string? key_search)
        {
            try
            {
                var query = from log in _context.ToolSupplyLogs
                            join tool in _context.Tools on log.ToolId equals tool.Id into toolJoin
                            from t in toolJoin.DefaultIfEmpty()
                            select new { log, t };

                if (!string.IsNullOrEmpty(type))
                {
                    query = query.Where(x => x.log.Type == type);
                }

                if (!string.IsNullOrEmpty(key_search))
                {
                    string search = key_search.ToLower();
                    query = query.Where(x =>
                        (x.t != null && x.t.ToolCode != null && x.t.ToolCode.ToLower().Contains(search)) ||
                        (x.t != null && x.t.ToolName != null && x.t.ToolName.ToLower().Contains(search)) ||
                        (x.log.WarehouseStaff != null && x.log.WarehouseStaff.ToLower().Contains(search))
                    );
                }

                var rawData = await query
                    .OrderByDescending(x => x.log.DateMonth)
                    .ThenByDescending(x => x.log.Id)
                    .ToListAsync();

                var result = rawData.Select(x => new
                {
                    Id = x.log.Id,
                    DateMonth = x.log.DateMonth.ToString("dd/MM/yyyy"),
                    Type = x.log.Type ?? "",
                    Qty = x.log.Qty,
                    WarehouseStaff = x.log.WarehouseStaff ?? "",
                    HandOverStaff = x.log.HandOverStaff ?? "",
                    Describe = x.log.Describe ?? "",
                    IntendedUse = x.log.IntendedUse ?? "",
                    Note = x.log.Note ?? "",
                    ToolCode = x.t != null ? (x.t.ToolCode ?? "") : "N/A",
                    ToolName = x.t != null ? (x.t.ToolName ?? "") : "N/A"
                }).ToList();

                return Json(result);
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [Authorize(Policy = "ViewToolSupplyLog")]
        public async Task<IActionResult> GetLogById(long id) => Json(await _context.ToolSupplyLogs.FindAsync(id));

        [Authorize(Policy = "EditToolSupplyLog")]
        [HttpPost]
        public async Task<IActionResult> SaveSupplyLog(ToolSupplyLog log)
        {
            var oldLog = await _context.ToolSupplyLogs.AsNoTracking().FirstOrDefaultAsync(x => x.Id == log.Id);
            if (oldLog == null)
            {
                return Json(new { success = false, message = "Không tìm thấy dữ liệu" });
            }

            var oldTool = await _context.Tools.FindAsync(oldLog.ToolId);
            if (oldTool != null)
            {
                if (oldLog.Type == "Nhập") oldTool.TotalImported -= oldLog.Qty;
                else if (oldLog.Type == "Hủy") oldTool.TotalScrapped -= oldLog.Qty;

                oldTool.AvailableQty = (oldTool.InitialQty + oldTool.TotalImported - oldTool.TotalScrapped) - oldTool.TotalIssued + oldTool.TotalReturned;
                oldTool.UpdatedDate = DateOnly.FromDateTime(DateTime.Now);
            }

            var newTool = await _context.Tools.FindAsync(log.ToolId);
            if (newTool == null)
            {
                return Json(new { success = false, message = "Vật tư không hợp lệ" });
            }

            if (log.Type == "Nhập") newTool.TotalImported += log.Qty;
            else if (log.Type == "Hủy") newTool.TotalScrapped += log.Qty;

            newTool.AvailableQty = (newTool.InitialQty + newTool.TotalImported - newTool.TotalScrapped) - newTool.TotalIssued + newTool.TotalReturned;
            newTool.UpdatedDate = DateOnly.FromDateTime(DateTime.Now);
            log.WarehouseStaff = oldLog.WarehouseStaff;
            log.UpdatedDate = DateOnly.FromDateTime(DateTime.Now);
            log.CreatedDate = oldLog.CreatedDate;

            _context.ToolSupplyLogs.Update(log);
            await _context.SaveChangesAsync();

            return Json(new { success = true });
        }

        [Authorize(Policy = "AddToolSupplyLog")]
        [HttpPost]
        public async Task<IActionResult> Import(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
                return BadRequest(new { Message = "Vui lòng chọn một file Excel." });

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var logList = new List<ToolSupplyLog>();
            var requiredFields = new[] { "Mã vật tư", "TNXT", "Số lượng", "Ngày" };

            try
            {
                using (var stream = excelFile.OpenReadStream())
                using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });

                    if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                        return BadRequest(new { Message = "File Excel không có dữ liệu." });

                    DataTable table = dataSet.Tables[0];
                    var missingColumns = requiredFields.Where(field => !table.Columns.Contains(field)).ToList();
                    if (missingColumns.Any())
                        return BadRequest(new { Message = $"Thiếu cột bắt buộc: {string.Join(", ", missingColumns)}" });

                    var toolsDict = await _context.Tools.ToDictionaryAsync(t => t.ToolCode.Trim(), t => t.Id);

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];
                        var rowIndex = i + 2;

                        var toolCode = row["Mã vật tư"]?.ToString()?.Trim();
                        var type = row["TNXT"]?.ToString()?.Trim();
                        var qtyStr = row["Số lượng"]?.ToString();
                        var dateStr = row["Ngày"]?.ToString();

                        if (string.IsNullOrEmpty(toolCode) || !toolsDict.ContainsKey(toolCode))
                            return BadRequest(new { Message = $"Lỗi dòng {rowIndex}: Mã vật tư '{toolCode}' không hợp lệ hoặc không tồn tại." });

                        if (string.IsNullOrEmpty(type))
                            return BadRequest(new { Message = $"Lỗi dòng {rowIndex}: Loại vật tư không được để trống." });

                        if (!int.TryParse(qtyStr, out int qty) || qty < 0)
                            return BadRequest(new { Message = $"Lỗi dòng {rowIndex}: Số lượng '{qtyStr}' phải là số nguyên dương." });

                        if (!DateTime.TryParse(dateStr, out DateTime parsedDate))
                            return BadRequest(new { Message = $"Lỗi dòng {rowIndex}: Định dạng ngày '{dateStr}' không hợp lệ." });

                        logList.Add(new ToolSupplyLog
                        {
                            ToolId = toolsDict[toolCode],
                            Type = type,
                            Qty = qty,
                            IntendedUse = table.Columns.Contains("Mục đích") ? row["Mục đích"]?.ToString() : null,
                            Describe = table.Columns.Contains("Mô tả") ? row["Mô tả"]?.ToString() : null,
                            WarehouseStaff = table.Columns.Contains("NV Kho") ? row["NV Kho"]?.ToString() : null,
                            HandOverStaff = table.Columns.Contains("NV bàn giao") ? row["NV bàn giao"]?.ToString() : null,
                            DateMonth = DateOnly.FromDateTime(parsedDate),
                            Note = table.Columns.Contains("Ghi chú") ? row["Ghi chú"]?.ToString() : null,
                            CreatedDate = DateOnly.FromDateTime(DateTime.Now)
                        });
                    }
                }

                if (logList.Any())
                {
                    await _context.ToolSupplyLogs.AddRangeAsync(logList);
                    await _context.SaveChangesAsync();
                }

                return Ok(new { Message = $"Import thành công {logList.Count} dòng dữ liệu." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { Message = "Lỗi hệ thống: " + ex.Message });
            }
        }

        [Authorize(Policy = "AddIssueReturnLog")]
        public async Task<IActionResult> DownloadTemplate()
        {
            var filePath = Path.Combine(_env.WebRootPath, "templates", "import_nhập/hủy vật tư.xlsx");
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Không tìm thấy file mẫu.");
            }
            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = "import_vật tư.xlsx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        [Authorize(Policy = "ViewToolSupplyLog")]
        public async Task<IActionResult> ExportToExcel()
        {
            var logs = await _context.ToolSupplyLogs
                .Join(_context.Tools,
                    log => log.ToolId,
                    tool => tool.Id,
                    (log, tool) => new
                    {
                        log.Id,
                        tool.ToolCode,
                        tool.ToolName,
                        tnxt = log.Type,
                        log.Qty,
                        log.IntendedUse,
                        log.Describe,
                        log.WarehouseStaff,
                        log.HandOverStaff,
                        log.DateMonth,
                        log.Note,
                        tool.Type,
                        tool.Unit,
                        tool.Location
                    })
                .OrderByDescending(x => x.DateMonth)
                .ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Data");

                worksheet.Cell(1, 1).Value = "THỐNG KÊ NHẬP/HỦY VẬT TƯ";
                worksheet.Range("A1:N1").Merge().Style.Font.Bold = true;
                worksheet.Range("A1:N1").Style.Font.FontSize = 16;
                worksheet.Range("A1:N1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var headers = new string[]
                {
                "STT", "Ngày", "TNXT", "Mã vật tư", "Tên vật tư", "Loại", "Đơn vị",
                "Số lượng", "Mục đích sử dụng", "Mô tả", "Vị trí", "NV Kho", "NV Nhận/Giao", "Ghi chú"
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
                foreach (var item in logs)
                {
                    worksheet.Cell(row, 1).Value = stt++;
                    worksheet.Cell(row, 2).Value = item.DateMonth.ToString("dd/MM/yyyy");
                    worksheet.Cell(row, 3).Value = item.tnxt;
                    worksheet.Cell(row, 4).Value = item.ToolCode;
                    worksheet.Cell(row, 5).Value = item.ToolName;
                    worksheet.Cell(row, 6).Value = item.Type;
                    worksheet.Cell(row, 7).Value = item.Unit;
                    worksheet.Cell(row, 8).Value = item.Qty;
                    worksheet.Cell(row, 9).Value = item.IntendedUse;
                    worksheet.Cell(row, 10).Value = item.Describe;
                    worksheet.Cell(row, 11).Value = item.Location;
                    worksheet.Cell(row, 12).Value = item.WarehouseStaff;
                    worksheet.Cell(row, 13).Value = item.HandOverStaff;
                    worksheet.Cell(row, 14).Value = item.Note;

                    worksheet.Range(row, 1, row, 14).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    row++;
                }

                worksheet.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    string fileName = $"Thống kê nhập/hủy vật tư.xlsx";

                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName
                    );
                }
            }
        }

        [Authorize(Policy = "DeleteToolSupplyLog")]
        public async Task<IActionResult> DeleteLog(long id)
        {
            var log = await _context.ToolSupplyLogs.FindAsync(id);
            if (log != null)
            {
                var tool = await _context.Tools.FindAsync(log.ToolId);
                if (log.Type == "Nhập") tool.TotalImported -= log.Qty;
                else tool.TotalScrapped -= log.Qty;

                tool.AvailableQty = (tool.InitialQty + tool.TotalImported - tool.TotalScrapped) - tool.TotalIssued + tool.TotalReturned;
                _context.ToolSupplyLogs.Remove(log);
                await _context.SaveChangesAsync();
            }
            return Json(new { success = true });
        }
    }
}
