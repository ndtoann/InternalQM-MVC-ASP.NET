using ClosedXML.Excel;
using ExcelDataReader.Log;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;

namespace Web_QM.Areas.Warehouse.Controllers
{
    [Area("Warehouse")]
    [Authorize]
    public class IssueReturnLogController : Controller
    {
        private readonly QMContext _context;

        public IssueReturnLogController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ViewIssueReturnLog")]
        public IActionResult Index()
        {
            return View();
        }

        [Authorize(Policy = "ViewIssueReturnLog")]
        public async Task<IActionResult> GetIssueReturnLogs(string? key_search, string? status, string? fromDate, string? toDate)
        {
            var query = from log in _context.IssueReturnLogs
                        join tool in _context.Tools on log.ToolId equals tool.Id into toolJoin
                        from t in toolJoin.DefaultIfEmpty()
                        select new { log, t };

            if (!string.IsNullOrEmpty(key_search))
            {
                string s = key_search.ToLower();
                query = query.Where(x => (x.t != null && x.t.ToolCode.ToLower().Contains(s)) ||
                                         (x.t != null && x.t.ToolName.ToLower().Contains(s)) ||
                                         (x.log.Machine.ToLower().Contains(s)));
            }

            if (!string.IsNullOrEmpty(status))
            {
                if (status == "pending")
                {
                    query = query.Where(x => x.log.ReturnDate == null);
                }
                else if (status == "completed")
                {
                    query = query.Where(x => x.log.ReturnDate != null);
                }
            }

            if (!string.IsNullOrEmpty(fromDate))
            {
                if (DateOnly.TryParse(fromDate, out DateOnly start))
                {
                    query = query.Where(x => x.log.IssuedDate >= start);
                }
            }

            if (!string.IsNullOrEmpty(toDate))
            {
                if (DateOnly.TryParse(toDate, out DateOnly end))
                {
                    query = query.Where(x => x.log.IssuedDate <= end);
                }
            }

            var data = await query.OrderByDescending(x => x.log.Id).ToListAsync();

            var result = data.Select(x => new {
                x.log.Id,
                x.log.ToolId,
                ToolCode = x.t?.ToolCode ?? "N/A",
                ToolName = x.t?.ToolName ?? "N/A",
                x.log.Machine,
                x.log.IntendedUse,
                IssuedDate = x.log.IssuedDate.ToString("dd/MM/yyyy"),
                x.log.IssuedQty,
                x.log.IssuedStaff,
                x.log.IssuedWarehouseStaff,
                ReturnDate = x.log.ReturnDate?.ToString("dd/MM/yyyy"),
                x.log.ReturnQty,
                x.log.ReturnStaff,
                x.log.ReturnWarehouseStaff,
                x.log.Note
            });

            return Json(result);
        }

        [Authorize(Policy = "AddIssueReturnLog")]
        public async Task<IActionResult> GetHoldingLogs(long toolId, string machine)
        {
            var logs = await _context.IssueReturnLogs
                .Where(x => x.ToolId == toolId && x.Machine == machine && x.IssuedQty > x.ReturnQty)
                .Select(x => new {
                    x.Id,
                    sortDate = x.IssuedDate,
                    IssuedDate = x.IssuedDate.ToString("dd/MM/yyyy"),
                    x.IssuedQty,
                    x.ReturnQty,
                    Debt = x.IssuedQty - x.ReturnQty,
                    x.IssuedStaff,
                    x.IssuedWarehouseStaff
                })
                .OrderByDescending(x => x.sortDate)
                .ToListAsync();
            return Json(logs);
        }

        [Authorize(Policy = "AddIssueReturnLog")]
        public async Task<IActionResult> GetMachines() => Json(await _context.Machines.Select(x => x.MachineCode).ToListAsync());

        [Authorize(Policy = "ViewIssueReturnLog")]
        public async Task<IActionResult> GetLogById(long id) => Json(await _context.IssueReturnLogs.FindAsync(id));

        [Authorize(Policy = "AddIssueReturnLog")]
        [HttpGet]
        public async Task<IActionResult> GetEmployeeSuggestions(string term)
        {
            var query = _context.Employees.AsQueryable();

            if (!string.IsNullOrEmpty(term))
            {
                query = query.Where(e => e.EmployeeCode.Contains(term) || e.EmployeeName.Contains(term));
            }

            var employees = await query
                .Select(e => new {
                    id = e.EmployeeCode,
                    text = e.EmployeeCode + " - " + e.EmployeeName
                })
                .ToListAsync();

            return Json(employees);
        }

        [Authorize(Policy = "AddIssueReturnLog")]
        [HttpPost]
        public async Task<IActionResult> SaveMultiIssue([FromBody] List<IssueReturnLog> models)
        {
            if (models == null || !models.Any())
            {
                return Json(new { success = false, message = "Không có dữ liệu" });
            }

            try
            {
                foreach (var log in models)
                {
                    var tool = await _context.Tools.FindAsync(log.ToolId);
                    if (tool == null) throw new Exception($"Vật tư không hợp lệ");

                    if (log.IssuedQty > tool.AvailableQty)
                        throw new Exception($"Vật tư {tool.ToolName} không đủ tồn (Còn: {tool.AvailableQty})");

                    tool.TotalIssued += log.IssuedQty;
                    tool.AvailableQty = (tool.InitialQty + tool.TotalImported - tool.TotalScrapped) - tool.TotalIssued + tool.TotalReturned;

                    log.CreatedDate = DateOnly.FromDateTime(DateTime.Now);
                    log.ReturnQty = 0;
                    log.ReturnDate = null;

                    _context.IssueReturnLogs.Add(log);
                }

                await _context.SaveChangesAsync();
                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [Authorize(Policy = "AddIssueReturnLog")]
        [HttpPost]
        public async Task<IActionResult> SaveIssueReturnLog(IssueReturnLog log)
        {
            try
            {
                if (log.IssuedQty < 0 || log.ReturnQty < 0)
                    return Json(new { success = false, message = "Số lượng không hợp lệ" });
                if (log.IssuedQty < log.ReturnQty)
                    return Json(new { success = false, message = "Số lượng trả không lớn hơn số lượng xuất" });

                if (log.ReturnDate == null) log.ReturnQty = 0;

                var existing = await _context.IssueReturnLogs.AsNoTracking().FirstOrDefaultAsync(x => x.Id == log.Id);
                if (existing == null) return Json(new { success = false, message = "Phiếu không tồn tại" });

                if (existing.ToolId != log.ToolId)
                {
                    var oldTool = await _context.Tools.FindAsync(existing.ToolId);
                    if (oldTool != null)
                    {
                        oldTool.TotalIssued -= existing.IssuedQty;
                        oldTool.TotalReturned -= existing.ReturnQty;
                        oldTool.AvailableQty = (oldTool.InitialQty + oldTool.TotalImported - oldTool.TotalScrapped) - oldTool.TotalIssued + oldTool.TotalReturned;
                    }

                    var newTool = await _context.Tools.FindAsync(log.ToolId);
                    if (newTool == null) return Json(new { success = false, message = "Vật tư không hợp lệ" });

                    if (log.IssuedQty > newTool.AvailableQty)
                        return Json(new { success = false, message = "Số lượng xuất vượt quá tồn kho" });

                    newTool.TotalIssued += log.IssuedQty;
                    newTool.TotalReturned += log.ReturnQty;
                    newTool.AvailableQty = (newTool.InitialQty + newTool.TotalImported - newTool.TotalScrapped) - newTool.TotalIssued + newTool.TotalReturned;
                }
                else
                {
                    var tool = await _context.Tools.FindAsync(log.ToolId);
                    if (tool == null) return Json(new { success = false, message = "Vật tư không hợp lệ" });

                    decimal qtyDifference = log.IssuedQty - existing.IssuedQty;
                    if (qtyDifference > tool.AvailableQty)
                        return Json(new { success = false, message = "Số lượng xuất vượt quá tồn kho" });

                    tool.TotalIssued = tool.TotalIssued - existing.IssuedQty + log.IssuedQty;
                    tool.TotalReturned = tool.TotalReturned - existing.ReturnQty + log.ReturnQty;
                    tool.AvailableQty = (tool.InitialQty + tool.TotalImported - tool.TotalScrapped) - tool.TotalIssued + tool.TotalReturned;
                }

                log.UpdatedDate = DateOnly.FromDateTime(DateTime.Now);
                _context.IssueReturnLogs.Update(log);
                await _context.SaveChangesAsync();
                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.InnerException?.Message ?? ex.Message });
            }
        }

        [Authorize(Policy = "ViewIssueReturnLog")]
        public async Task<IActionResult> ExportToExcel()
        {
            var logs = await _context.IssueReturnLogs
                .Join(_context.Tools,
                    log => log.ToolId,
                    tool => tool.Id,
                    (log, tool) => new
                    {
                        log.IssuedDate,
                        tool.ToolCode,
                        tool.ToolName,
                        tool.Type,
                        tool.Unit,
                        log.IssuedQty,
                        log.Machine,
                        log.IntendedUse,
                        tool.Location,
                        log.Note,
                        log.IssuedWarehouseStaff,
                        log.ReturnWarehouseStaff,
                        log.ReturnStaff,
                        log.IssuedStaff,
                        log.ReturnDate,
                        log.ReturnQty,
                    })
                .OrderByDescending(x => x.IssuedDate)
                .ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Data");

                worksheet.Cell(1, 1).Value = "Xuất trả vật tư hàng ngày";
                worksheet.Range("A1:P1").Merge().Style.Font.Bold = true;
                worksheet.Range("A1:P1").Style.Font.FontSize = 16;
                worksheet.Range("A1:P1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var headers = new string[]
                {
                "Ngày", "Mã vật tư", "Tên vật tư", "Loại", "Đơn vị", "SL xuất",
                "Máy", "Mục đích sử dụng", "NV kho", "NV mượn",
                "Ngày trả", "SL trả", "NV Kho", "NV trả", "Vị trí", "Ghi chú"
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
                foreach (var item in logs)
                {
                    worksheet.Cell(row, 1).Value = item.IssuedDate.ToString("dd/MM/yyyy");
                    worksheet.Cell(row, 2).Value = item.ToolCode;
                    worksheet.Cell(row, 3).Value = item.ToolName;
                    worksheet.Cell(row, 4).Value = item.Type;
                    worksheet.Cell(row, 5).Value = item.Unit;
                    worksheet.Cell(row, 6).Value = item.IssuedQty;
                    worksheet.Cell(row, 7).Value = item.Machine;
                    worksheet.Cell(row, 8).Value = item.IntendedUse;
                    worksheet.Cell(row, 9).Value = item.IssuedWarehouseStaff;
                    worksheet.Cell(row, 10).Value = item.IssuedStaff;
                    worksheet.Cell(row, 11).Value = item.ReturnDate?.ToString("dd/MM/yyyy") ?? "";
                    worksheet.Cell(row, 12).Value = item.ReturnQty;
                    worksheet.Cell(row, 13).Value = item.ReturnWarehouseStaff;
                    worksheet.Cell(row, 14).Value = item.ReturnStaff;
                    worksheet.Cell(row, 15).Value = item.Location;
                    worksheet.Cell(row, 16).Value = item.Note;

                    worksheet.Range(row, 1, row, 16).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    row++;
                }

                worksheet.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"Xuất trả vật tư hàng ngày.xlsx");
                }
            }
        }

        [Authorize(Policy = "DeleteIssueReturnLog")]
        public async Task<IActionResult> DeleteLog(long id)
        {
            try
            {
                var log = await _context.IssueReturnLogs.FindAsync(id);
                if (log != null)
                {
                    var tool = await _context.Tools.FindAsync(log.ToolId);
                    if (tool != null)
                    {
                        tool.TotalIssued -= log.IssuedQty;
                        tool.TotalReturned -= log.ReturnQty;
                        tool.AvailableQty = (tool.InitialQty + tool.TotalImported - tool.TotalScrapped) - tool.TotalIssued + tool.TotalReturned;
                    }
                    _context.IssueReturnLogs.Remove(log);
                    await _context.SaveChangesAsync();
                }
                return Json(new { success = true });
            }
            catch (Exception ex) { return Json(new { success = false, message = ex.Message }); }
        }
    }
}
