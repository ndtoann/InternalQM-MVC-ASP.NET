using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Data;
using System.Text.RegularExpressions;
using Web_QM.Models;

namespace Web_QM.Areas.Production.Controllers
{
    [Area("Production")]
    public class ProductionProcessessController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public ProductionProcessessController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewProductionProcessess")]
        public async Task<IActionResult> Index()
        {
            return View();
        }

        [Authorize(Policy = "ViewProductionProcessess")]
        public async Task<IActionResult> GetList(string key_search, string status)
        {
            var query = _context.ProductionProcessess.AsQueryable();

            if (!string.IsNullOrEmpty(key_search))
            {
                key_search = key_search.ToLower();
                query = query.Where(x => x.PartName.ToLower().Contains(key_search)
                                      || (x.Material != null && x.Material.ToLower().Contains(key_search)));
            }

            if (status == "new")
            {
                var latestIds = await query.GroupBy(x => x.PartName)
                                           .Select(g => g.Max(x => x.Id))
                                           .ToListAsync();
                query = query.Where(x => latestIds.Contains(x.Id));
            }

            var result = await query.OrderByDescending(x => x.Id).Take(200).ToListAsync();
            return Json(result);
        }

        [Authorize(Policy = "ViewProductionProcessess")]
        public async Task<IActionResult> GetDetails(long id)
        {
            var header = await _context.ProductionProcessess.FindAsync(id);
            var steps = await _context.ProcessSteps
                .Where(x => x.ProductionProcessId == id)
                .OrderBy(x => x.StepNumber)
                .ToListAsync();

            var result = new
            {
                header = new
                {
                    header.Id,
                    header.PartName,
                    header.WorkpieceSize,
                    header.Material,
                    header.Picture,
                    header.Version,
                    header.CreatedBy,
                    CreatedDate = header.CreatedDate.ToString("dd/MM/yyyy"),
                    header.UpdatedBy,
                    UpdatedDate = header.UpdatedDate?.ToString("dd/MM/yyyy"),
                    header.Note,
                    header.ModifiedContent
                },
                steps = steps.Select(s => new {
                    s.Id,
                    s.StepNumber,
                    s.Department,
                    s.Content,
                    s.Fixture,
                    EstimatedTime = s.EstimatedTime.ToString("N1"),
                    s.Picture,
                    s.Note
                })
            };
            return Json(result);
        }

        [Authorize(Policy = "AddProductionProcessess")]
        [HttpPost]
        public async Task<IActionResult> Save(ProductionProcessess header, List<ProcessStep> steps, bool createNewVersion)
        {
            try
            {
                var currentUser = User.FindFirst("EmployeeCode")?.Value + "-" + User.FindFirst("EmployeeName")?.Value;
                if (currentUser == "-")
                {
                    return Json(new { success = false, message = "Vui lòng đăng nhập lại" });
                }

                if (string.IsNullOrWhiteSpace(header.PartName))
                    return Json(new { success = false, message = "Tên chi tiết trống." });

                if (steps == null || !steps.Any())
                    return Json(new { success = false, message = "Thiếu nguyên công." });

                foreach (var s in steps)
                {
                    if (s.EstimatedTime < 0)
                    {
                        return Json(new { success = false, message = $"Thời gian gia công không hợp lệ." });
                    }
                }

                var headerFile = Request.Form.Files.GetFile("HeaderFile");

                if (header.Id == 0 || createNewVersion)
                {
                    if (header.Id == 0 && !createNewVersion)
                    {
                        var isExist = await _context.ProductionProcessess.AnyAsync(x => x.PartName == header.PartName);
                        if (isExist)
                        {
                            return Json(new { success = false, message = $"Chi tiết '{header.PartName}' đã có quy trình." });
                        }
                    }

                    var oldId = header.Id;
                    if (createNewVersion)
                    {
                        var oldData = await _context.ProductionProcessess.AsNoTracking().FirstOrDefaultAsync(x => x.Id == oldId);
                        var maxV = await _context.ProductionProcessess.Where(x => x.PartName == header.PartName).MaxAsync(x => (int?)x.Version) ?? 0;
                        header.Version = maxV + 1;
                        header.Picture = headerFile != null ? await SaveFile(headerFile) : await CopyFile(oldData?.Picture);
                    }
                    else
                    {
                        header.Version = 1;
                        if (headerFile != null) header.Picture = await SaveFile(headerFile);
                    }

                    header.Id = 0;
                    header.CreatedBy = currentUser;
                    header.CreatedDate = DateOnly.FromDateTime(DateTime.Now);
                    header.UpdatedBy = null;
                    header.UpdatedDate = null;

                    _context.ProductionProcessess.Add(header);
                    await _context.SaveChangesAsync();
                }
                else
                {
                    var existing = await _context.ProductionProcessess.FindAsync(header.Id);
                    if (existing == null) return Json(new { success = false, message = "Không thấy dữ liệu." });

                    existing.PartName = header.PartName;
                    existing.WorkpieceSize = header.WorkpieceSize;
                    existing.Material = header.Material;
                    existing.Note = header.Note;
                    existing.ModifiedContent = header.ModifiedContent;
                    existing.UpdatedBy = currentUser;
                    existing.UpdatedDate = DateOnly.FromDateTime(DateTime.Now);

                    if (headerFile != null)
                    {
                        DeleteFile(existing.Picture);
                        existing.Picture = await SaveFile(headerFile);
                    }

                    var oldSteps = await _context.ProcessSteps.Where(x => x.ProductionProcessId == header.Id).ToListAsync();
                    var currentPics = steps.Select(s => s.Picture).Where(p => !string.IsNullOrEmpty(p)).ToList();

                    foreach (var os in oldSteps)
                    {
                        if (!string.IsNullOrEmpty(os.Picture) && !currentPics.Contains(os.Picture)) DeleteFile(os.Picture);
                    }

                    _context.ProcessSteps.RemoveRange(oldSteps);
                    await _context.SaveChangesAsync();
                }

                for (int i = 0; i < steps.Count; i++)
                {
                    var s = steps[i];
                    var f = Request.Form.Files.GetFile($"StepFile_{i}");
                    var step = new ProcessStep
                    {
                        ProductionProcessId = header.Id,
                        StepNumber = i + 1,
                        Department = s.Department,
                        QtyPerSet = string.IsNullOrWhiteSpace(s.QtyPerSet) ? "1" : s.QtyPerSet,
                        Content = s.Content,
                        Fixture = s.Fixture,
                        EstimatedTime = s.EstimatedTime,
                        Note = s.Note
                    };

                    if (f != null) step.Picture = await SaveFile(f);
                    else if (createNewVersion) step.Picture = await CopyFile(s.Picture);
                    else step.Picture = s.Picture;

                    _context.ProcessSteps.Add(step);
                }

                await _context.SaveChangesAsync();
                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.InnerException?.Message ?? ex.Message });
            }
        }

        [Authorize(Policy = "DeleteProductionProcessess")]
        [HttpPost]
        public async Task<IActionResult> Delete(long id)
        {
            try
            {
                var header = await _context.ProductionProcessess.FindAsync(id);
                if (header == null) return Json(new { success = false, message = "Không tìm thấy dữ liệu." });

                DeleteFile(header.Picture);
                var steps = await _context.ProcessSteps.Where(x => x.ProductionProcessId == id).ToListAsync();
                foreach (var s in steps) DeleteFile(s.Picture);

                _context.ProcessSteps.RemoveRange(steps);
                _context.ProductionProcessess.Remove(header);
                await _context.SaveChangesAsync();
                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [Authorize(Policy = "ViewProductionProcessess")]
        public async Task<IActionResult> ExportExcel(int id)
        {
            try
            {
                var header = await _context.ProductionProcessess.FirstOrDefaultAsync(x => x.Id == id);
                if (header == null) return Json(new { success = false, message = "Không tìm thấy dữ liệu" });

                var steps = await _context.ProcessSteps.Where(x => x.ProductionProcessId == id).OrderBy(x => x.StepNumber).ToListAsync();

                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Quy trình");

                    ws.Cell("A1").Value = "QUY TRÌNH GIA CÔNG";
                    ws.Range("A1:H1").Merge().Style.Font.SetBold().Font.SetFontSize(16).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                    ws.Cell("A2").Value = "Tên chi tiết:";
                    ws.Cell("B2").Value = header.PartName;
                    ws.Cell("B2").Style.Font.SetBold();
                    ws.Cell("A3").Value = "Kích thước phôi:";
                    ws.Cell("B3").Value = header.WorkpieceSize;
                    ws.Cell("B3").Style.Font.SetBold();
                    ws.Cell("A4").Value = "Vật liệu:";
                    ws.Cell("B4").Value = header.Material;
                    ws.Cell("A5").Value = "Ghi chú:";
                    ws.Cell("B5").Value = header.Note;

                    ws.Cell("F2").Value = "Ngày lập:";
                    ws.Cell("G2").Value = header.CreatedDate.ToString("dd/MM/yyyy");
                    ws.Cell("H2").Value = header.CreatedBy;
                    ws.Cell("F3").Value = "Phiên bản:";
                    ws.Cell("G3").Value = "V" + header.Version;
                    ws.Cell("F4").Value = "Ngày sửa:";
                    ws.Cell("G4").Value = header.UpdatedDate?.ToString("dd/MM/yyyy") ?? "";
                    ws.Cell("H4").Value = header.UpdatedBy ?? "";
                    ws.Cell("F5").Value = "Nội dung sửa đổi:";
                    ws.Cell("G5").Value = header.ModifiedContent;

                    ws.Row(6).Height = 120;
                    var detailImgRange = ws.Range("A6:H6");
                    detailImgRange.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center).Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                    if (!string.IsNullOrEmpty(header.Picture))
                    {
                        string fullPath = Path.Combine(_env.WebRootPath, header.Picture.TrimStart('/'));
                        if (System.IO.File.Exists(fullPath))
                        {
                            ws.AddPicture(fullPath)
                              .MoveTo(ws.Cell(6, 1))
                              .WithSize(150, 150);
                        }
                    }

                    ws.Row(7).Height = 25;
                    ws.Cell(7, 1).Value = "Ảnh mô tả";
                    ws.Cell(7, 2).Value = "Nguyên công";
                    ws.Cell(7, 3).Value = "Bộ phận";
                    ws.Cell(7, 4).Value = "Nội dung gia công";
                    ws.Cell(7, 5).Value = "Đồ gá";
                    ws.Cell(7, 6).Value = "Thời gian gia công";
                    ws.Cell(7, 7).Value = "SL/1 lần gá";
                    ws.Cell(7, 8).Value = "Ghi chú";

                    var headerRange = ws.Range("A7:H7");
                    headerRange.Style
                        .Fill.SetBackgroundColor(XLColor.LightGray)
                        .Font.SetBold()
                        .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                        .Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                    for (int i = 0; i < steps.Count; i++)
                    {
                        int row = i + 8;
                        ws.Row(row).Height = 120;

                        if (!string.IsNullOrEmpty(steps[i].Picture))
                        {
                            string stepImgPath = Path.Combine(_env.WebRootPath, steps[i].Picture.TrimStart('/'));
                            if (System.IO.File.Exists(stepImgPath))
                            {
                                ws.AddPicture(stepImgPath)
                                  .MoveTo(ws.Cell(row, 1))
                                  .WithSize(150, 150);
                            }
                        }

                        ws.Cell(row, 2).Value = $"{i + 1}/{steps.Count}";
                        ws.Cell(row, 3).Value = steps[i].Department;
                        ws.Cell(row, 4).Value = steps[i].Content;
                        ws.Cell(row, 5).Value = steps[i].Fixture;
                        ws.Cell(row, 6).Value = steps[i].EstimatedTime + " phút";
                        ws.Cell(row, 7).Value = steps[i].QtyPerSet;
                        ws.Cell(row, 8).Value = steps[i].Note;

                        ws.Cell(row, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(row, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(row, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(row, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Cell(row, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        ws.Range(row, 1, row, 8).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    }

                    ws.Column(1).Width = 22;
                    ws.Columns(2, 9).AdjustToContents();

                    ws.RangeUsed().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin).Border.SetInsideBorder(XLBorderStyleValues.Thin);

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        string safeFileName = Regex.Replace(header.PartName, @"[\\/:*?""<>|]", "_");
                        string fileName = $"{safeFileName}.xlsx";
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        private async Task<string> SaveFile(IFormFile file)
        {
            var folder = Path.Combine(_env.WebRootPath, "imgs", "production_processess");
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            var fileName = Guid.NewGuid().ToString() + Path.GetExtension(file.FileName);
            var path = Path.Combine(folder, fileName);
            using (var stream = new FileStream(path, FileMode.Create)) await file.CopyToAsync(stream);
            return "/imgs/production_processess/" + fileName;
        }

        private async Task<string?> CopyFile(string? relPath)
        {
            if (string.IsNullOrEmpty(relPath)) return null;
            var oldPath = Path.Combine(_env.WebRootPath, relPath.TrimStart('/'));
            if (!System.IO.File.Exists(oldPath)) return null;
            var folder = Path.Combine(_env.WebRootPath, "imgs", "production_processess");
            var newName = Guid.NewGuid().ToString() + Path.GetExtension(oldPath);
            var newPath = Path.Combine(folder, newName);
            System.IO.File.Copy(oldPath, newPath);
            return "/imgs/production_processess/" + newName;
        }

        private void DeleteFile(string? relPath)
        {
            if (string.IsNullOrEmpty(relPath)) return;
            var path = Path.Combine(_env.WebRootPath, relPath.TrimStart('/'));
            if (System.IO.File.Exists(path)) System.IO.File.Delete(path);
        }

        [Authorize(Policy = "ViewProductionProcessess")]
        public async Task<IActionResult> Print(int id)
        {
            var header = await _context.ProductionProcessess.FirstOrDefaultAsync(x => x.Id == id);
            if (header == null) return NotFound();

            var steps = await _context.ProcessSteps
                .Where(x => x.ProductionProcessId == id)
                .OrderBy(x => x.StepNumber)
                .ToListAsync();

            ViewBag.Steps = steps;
            return View(header);
        }

        [Authorize(Policy = "AddProductionProcessess")]
        [HttpPost]
        public async Task<IActionResult> Import(IFormFile excelFile)
        {
            var currentUser = User.FindFirst("EmployeeCode")?.Value + "-" + User.FindFirst("EmployeeName")?.Value;
            if (currentUser == "-")
            {
                return BadRequest(new { Message = "Vui lòng đăng nhập lại" });
            }
            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest(new { Message = "Vui lòng chọn một file Excel." });
            }

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var requiredFields = new[] { "Tên chi tiết", "Kích thước phôi", "Vật liệu", "Nguyên công", "Bộ phận", "TGGC" };
            var tempFlatData = new List<dynamic>();

            try
            {
                using (var stream = excelFile.OpenReadStream())
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });

                    if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                    {
                        return BadRequest(new { Message = "File Excel không có dữ liệu." });
                    }

                    DataTable table = dataSet.Tables[0];
                    var missingColumns = requiredFields.Where(field => !table.Columns.Contains(field)).ToList();
                    if (missingColumns.Any())
                    {
                        return BadRequest(new { Message = $"File Excel thiếu các cột: {string.Join(", ", missingColumns)}" });
                    }

                    var duplicateCheckInFile = new HashSet<string>();

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        var row = table.Rows[i];
                        var rowIndex = i + 2;

                        foreach (var field in requiredFields)
                        {
                            if (row[field] == DBNull.Value || string.IsNullOrWhiteSpace(row[field].ToString()))
                            {
                                return BadRequest(new { Message = $"Lỗi dòng {rowIndex}: Thiếu dữ liệu cột '{field}'." });
                            }
                        }

                        var partName = row["Tên chi tiết"].ToString().Trim();
                        var stepNumber = int.Parse(row["Nguyên công"].ToString());
                        var duplicateKey = $"{partName}_{stepNumber}";

                        if (duplicateCheckInFile.Contains(duplicateKey))
                        {
                            return BadRequest(new { Message = $"Lỗi dòng {rowIndex}: Trùng lặp tên chi tiết '{partName}' và nguyên công '{stepNumber}' trong file." });
                        }
                        duplicateCheckInFile.Add(duplicateKey);

                        tempFlatData.Add(new
                        {
                            PartName = partName,
                            WorkpieceSize = row["Kích thước phôi"].ToString().Trim(),
                            Material = row["Vật liệu"].ToString().Trim(),
                            ProcessNote = table.Columns.Contains("Ghi chú") ? row["Ghi chú"]?.ToString() : null,
                            ModifiedContent = table.Columns.Contains("Nội dung sửa đổi") ? row["Nội dung sửa đổi"]?.ToString() : null,
                            StepNote = table.Columns.Contains("Chú ý") ? row["Chú ý"]?.ToString() : null,
                            StepNumber = stepNumber,
                            Department = row["Bộ phận"].ToString().Trim(),
                            Content = table.Columns.Contains("Nội dung gia công") ? row["Nội dung gia công"]?.ToString() : null,
                            EstimatedTime = decimal.Parse(row["TGGC"].ToString()),
                            Fixture = table.Columns.Contains("Đồ gá") ? row["Đồ gá"]?.ToString() : null,
                            QtyPerSet = table.Columns.Contains("SL/1 lần gá") ? row["SL/1 lần gá"]?.ToString() : "1",
                            RowIndex = rowIndex
                        });
                    }
                }

                var partNamesInFile = tempFlatData.Select(x => (string)x.PartName).Distinct().ToList();
                var existingPartNames = await _context.ProductionProcessess
                    .Where(p => partNamesInFile.Contains(p.PartName))
                    .Select(p => p.PartName)
                    .ToListAsync();

                if (existingPartNames.Any())
                {
                    return BadRequest(new { Message = $"Chi tiết đã có quy trình: {string.Join(", ", existingPartNames)}" });
                }

                var groupedData = tempFlatData.GroupBy(x => x.PartName).ToList();

                foreach (var group in groupedData)
                {
                    var partName = group.Key;
                    var steps = group.OrderBy(x => x.StepNumber).ToList();

                    for (int i = 0; i < steps.Count; i++)
                    {
                        int expectedStep = i + 1;
                        if (steps[i].StepNumber != expectedStep)
                        {
                            return BadRequest(new { Message = $"Lỗi chi tiết '{partName}': Thứ tự nguyên công không đúng." });
                        }
                    }

                    var firstRow = steps.First();
                    var process = new ProductionProcessess
                    {
                        PartName = firstRow.PartName,
                        WorkpieceSize = firstRow.WorkpieceSize,
                        Material = firstRow.Material,
                        Note = firstRow.ProcessNote,
                        ModifiedContent = firstRow.ModifiedContent,
                        CreatedDate = DateOnly.FromDateTime(DateTime.Now),
                        Version = 1,
                        CreatedBy = currentUser
                    };

                    _context.ProductionProcessess.Add(process);
                    await _context.SaveChangesAsync();

                    var processSteps = steps.Select(s => new ProcessStep
                    {
                        ProductionProcessId = process.Id,
                        StepNumber = s.StepNumber,
                        Department = s.Department,
                        Content = s.Content,
                        EstimatedTime = s.EstimatedTime,
                        Fixture = s.Fixture,
                        QtyPerSet = s.QtyPerSet,
                        Note = s.StepNote
                    }).ToList();

                    _context.ProcessSteps.AddRange(processSteps);
                    await _context.SaveChangesAsync();
                }

                return Ok(new { Message = $"Import thành công {groupedData.Count} chi tiết." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { Message = $"Lỗi hệ thống: {ex.Message}" });
            }
        }

        [Authorize(Policy = "AddProductionProcessess")]
        public async Task<IActionResult> DownloadTemplate()
        {
            var filePath = Path.Combine(_env.WebRootPath, "templates", "import_quy trình.xlsx");
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Không tìm thấy file mẫu.");
            }
            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = "import_quy trình.xlsx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }
}
