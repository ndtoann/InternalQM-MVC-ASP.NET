using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;

namespace Web_QM.Areas.Warehouse.Controllers
{
    [Area("Warehouse")]
    [Authorize]
    public class ToolSupplyLogController : Controller
    {
        private readonly QMContext _context;

        public ToolSupplyLogController(QMContext context)
        {
            _context = context;
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

        [Authorize(Policy = "AddIssueReturnLog")]
        [HttpPost]
        public async Task<IActionResult> SaveLog(ToolSupplyLog log)
        {
            var tool = await _context.Tools.FindAsync(log.ToolId);
            var old = await _context.ToolSupplyLogs.AsNoTracking().FirstOrDefaultAsync(x => x.Id == log.Id);

            if (old != null)
            {
                if (old.Type == "Nhập") tool.TotalImported -= old.Qty;
                else tool.TotalScrapped -= old.Qty;

                if (log.Type == "Nhập") tool.TotalImported += log.Qty;
                else tool.TotalScrapped += log.Qty;

                _context.ToolSupplyLogs.Update(log);
                tool.AvailableQty = (tool.InitialQty + tool.TotalImported - tool.TotalScrapped) - tool.TotalIssued + tool.TotalReturned;
                await _context.SaveChangesAsync();
            }
            return Json(new { success = true });
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
