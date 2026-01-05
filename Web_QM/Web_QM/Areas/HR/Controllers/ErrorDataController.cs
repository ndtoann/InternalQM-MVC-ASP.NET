using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;
using System.Data;
using Web_QM.Models;
using Web_QM.Models.ViewModels;

namespace Web_QM.Areas.HR.Controllers
{
    [Area("HR")]
    public class ErrorDataController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public ErrorDataController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewProductionDefect")]
        public async Task<IActionResult> Index(string key, DateTime? startDate = null, DateTime? endDate = null)
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

            ViewBag.KeySearch = key;
            ViewBag.StartDate = startDate.Value.ToString("yyyy-MM-dd");
            ViewBag.EndDate = endDate.Value.Date.ToString("yyyy-MM-dd");
            return View(res);
        }

        [Authorize(Policy = "ViewProductionDefect")]
        public async Task<IActionResult> GetDetail(long id)
        {
            var errorData = await _context.ErrorDatas
                                    .FirstOrDefaultAsync(e => e.Id == id);

            if (errorData == null)
            {
                return NotFound();
            }

            return Json(errorData);
        }

        [Authorize(Policy = "AddProductionDefect")]
        public async Task<IActionResult> Add()
        {
            var employees = await _context.Employees.AsNoTracking().OrderBy(o => o.EmployeeCode).ToListAsync();
            var employeeList = employees.Select(e => new
            {
                Value = e.EmployeeCode,
                Text = e.EmployeeCode + " - " + e.EmployeeName
            });
            ViewBag.ListEmployee = new SelectList(employeeList, "Value", "Text");

            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewBag.ListDepartment = departmentsList;
            return View();
        }

        [Authorize(Policy = "ViewAddProductionDefectKaizen")]
        [HttpPost]
        public async Task<IActionResult> Add(ErrorData errorData)
        {
            var employees = await _context.Employees.AsNoTracking().OrderBy(o => o.EmployeeCode).ToListAsync();
            var employeeList = employees.Select(e => new
            {
                Value = e.EmployeeCode,
                Text = e.EmployeeCode + " - " + e.EmployeeName
            });
            ViewBag.ListEmployee = new SelectList(employeeList, "Value", "Text");

            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewBag.ListDepartment = departmentsList;

            if (!ModelState.IsValid)
            {
                return View(errorData);
            }
            try
            {
                _context.ErrorDatas.Add(errorData);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction("Index");
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(errorData);
            }
        }

        [Authorize(Policy = "EditProductionDefect")]
        public async Task<IActionResult> Edit(long id)
        {
            var errorData = await _context.ErrorDatas.FindAsync(id);
            if (errorData == null)
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

            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewBag.ListDepartment = departmentsList;
            return View(errorData);
        }

        [Authorize(Policy = "EditProductionDefect")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, ErrorData errorData)
        {
            if (id != errorData.Id)
            {
                return NotFound();
            }
            var errorDataToEdit = await _context.ErrorDatas.AsNoTracking().FirstOrDefaultAsync(e => e.Id == id);
            if (errorDataToEdit == null)
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

            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewBag.ListDepartment = departmentsList;
            if (!ModelState.IsValid)
            {
                return View(errorData);
            }
            try
            {
                _context.ErrorDatas.Update(errorData);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction("Index");
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(errorData);
            }
        }

        [Authorize(Policy = "DeleteProductionDefect")]
        public async Task<IActionResult> Delete(long id)
        {
            var errorData = await _context.ErrorDatas.FindAsync(id);
            if (errorData == null)
            {
                return NotFound();
            }
            try
            {
                _context.ErrorDatas.Remove(errorData);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa báo lỗi!";
                return RedirectToAction("Index");
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
                return RedirectToAction("Index");
            }
        }

        [Authorize(Policy = "AddProductionDefect")]
        [HttpPost]
        public async Task<IActionResult> Import(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest(new { Message = "Vui lòng chọn một file Excel." });
            }

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var errorDataList = new List<ErrorData>();
            var requiredFields = new[] {
            "OrderNo", "PartName", "QtyOrder", "QtyNG", "Ngày Tháng",
            "Phát hiện lỗi", "Dạng lỗi", "Nội dung", "NCC"
        };

            try
            {
                using (var stream = excelFile.OpenReadStream())
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
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

                        var missingColumns = requiredFields.Where(field => !table.Columns.Contains(field)).ToList();
                        if (missingColumns.Any())
                        {
                            return BadRequest(new { Message = $"File Excel thiếu các cột bắt buộc: {string.Join(", ", missingColumns)}" });
                        }

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
                                DateOnly dateMonth;
                                object dateValue = row["Ngày Tháng"];

                                if (dateValue is DateTime dt)
                                {
                                    dateMonth = DateOnly.FromDateTime(dt);
                                }
                                else if (double.TryParse(dateValue.ToString(), out double doubleDate) && doubleDate > 1)
                                {
                                    dateMonth = DateOnly.FromDateTime(DateTime.FromOADate(doubleDate));
                                }
                                else if (DateTime.TryParse(dateValue.ToString(), out DateTime parsedDt))
                                {
                                    dateMonth = DateOnly.FromDateTime(parsedDt);
                                }
                                else
                                {
                                    throw new FormatException("Định dạng Ngày/tháng không hợp lệ.");
                                }

                                if (!int.TryParse(row["QtyOrder"].ToString(), out int qtyOrder))
                                {
                                    throw new FormatException("QtyOrder phải là số nguyên.");
                                }
                                if (!int.TryParse(row["QtyNG"].ToString(), out int qtyNG))
                                {
                                    throw new FormatException("QtyNG phải là số nguyên.");
                                }

                                DateOnly? errorCompletionDate = null;
                                if (table.Columns.Contains("Ngày hoàn thành giấy báo lỗi") && row["Ngày hoàn thành giấy báo lỗi"] != DBNull.Value)
                                {
                                    object completionValue = row["Ngày hoàn thành giấy báo lỗi"];
                                    DateTime tempDt;

                                    if (completionValue is DateTime cDt)
                                    {
                                        errorCompletionDate = DateOnly.FromDateTime(cDt);
                                    }
                                    else if (double.TryParse(completionValue.ToString(), out double cDoubleDate) && cDoubleDate > 1)
                                    {
                                        errorCompletionDate = DateOnly.FromDateTime(DateTime.FromOADate(cDoubleDate));
                                    }
                                    else if (DateTime.TryParse(completionValue.ToString(), out tempDt))
                                    {
                                        errorCompletionDate = DateOnly.FromDateTime(tempDt);
                                    }
                                }

                                var errorData = new ErrorData
                                {
                                    OrderNo = row["OrderNo"].ToString().Trim(),
                                    PartName = row["PartName"].ToString().Trim(),
                                    QtyOrder = qtyOrder,
                                    QtyNG = qtyNG,
                                    DateMonth = dateMonth,
                                    ErrorDetected = row["Phát hiện lỗi"].ToString().Trim(),
                                    ErrorType = row["Dạng lỗi"].ToString().Trim(),
                                    ErrorContent = row["Nội dung"].ToString().Trim(),
                                    NCC = row["NCC"].ToString().Trim(),

                                    ErrorCause = row.Table.Columns.Contains("Nguyên nhân lỗi") && row["Nguyên nhân lỗi"] != DBNull.Value ? row["Nguyên nhân lỗi"].ToString().Trim() : null,
                                    ToleranceAssessment = row.Table.Columns.Contains("Nhận định dung sai") && row["Nhận định dung sai"] != DBNull.Value ? row["Nhận định dung sai"].ToString().Trim() : null,
                                    Reason = row.Table.Columns.Contains("Nguyên nhân") && row["Nguyên nhân"] != DBNull.Value ? row["Nguyên nhân"].ToString().Trim() : null,
                                    Countermeasure = row.Table.Columns.Contains("Đối sách") && row["Đối sách"] != DBNull.Value ? row["Đối sách"].ToString().Trim() : null,
                                    EmployeeCode = row.Table.Columns.Contains("Mã nhân viên") && row["Mã nhân viên"] != DBNull.Value ? row["Mã nhân viên"].ToString().Trim() : null,
                                    Department = row.Table.Columns.Contains("Bộ phận") && row["Bộ phận"] != DBNull.Value ? row["Bộ phận"].ToString().Trim() : null,
                                    ErrorCompletionDate = errorCompletionDate,
                                    RemedialMeasures = row.Table.Columns.Contains("Biện pháp khắc phục") && row["Biện pháp khắc phục"] != DBNull.Value ? row["Biện pháp khắc phục"].ToString().Trim() : null,
                                    Note = row.Table.Columns.Contains("Ghi chú") && row["Ghi chú"] != DBNull.Value ? row["Ghi chú"].ToString().Trim() : null,
                                    TimeWriteError = row.Table.Columns.Contains("Thời gian viết giấy lỗi") && row["Thời gian viết giấy lỗi"] != DBNull.Value ? row["Thời gian viết giấy lỗi"].ToString().Trim() : null,
                                    ReviewNnds = row.Table.Columns.Contains("Rà soát việc thực hiện NNDS") && row["Rà soát việc thực hiện NNDS"] != DBNull.Value ? row["Rà soát việc thực hiện NNDS"].ToString().Trim() : null,
                                };

                                var validationContext = new ValidationContext(errorData, serviceProvider: null, items: null);
                                var validationResults = new List<ValidationResult>();

                                if (!Validator.TryValidateObject(errorData, validationContext, validationResults, true))
                                {
                                    var errors = string.Join("; ", validationResults.Select(r => r.ErrorMessage));
                                    return BadRequest(new { Message = $"Lỗi tại dòng {rowIndex}: Lỗi dữ liệu. Chi tiết: {errors}" });
                                }

                                errorDataList.Add(errorData);
                            }
                            catch (FormatException ex)
                            {
                                return BadRequest(new { Message = $"Lỗi tại dòng {rowIndex}: Lỗi chuyển đổi dữ liệu. Chi tiết: {ex.Message}" });
                            }
                            catch (Exception ex)
                            {
                                return BadRequest(new { Message = $"Lỗi tại dòng {rowIndex}: Lỗi không xác định khi xử lý dữ liệu. Chi tiết: {ex.Message}" });
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
                if (errorDataList.Any())
                {
                    await _context.ErrorDatas.AddRangeAsync(errorDataList);
                    await _context.SaveChangesAsync();
                }

                return Ok(new { Message = $"Đã import và lưu thành công {errorDataList.Count} bản ghi lỗi sản xuất." });
            }
            catch (DbUpdateException ex)
            {
                return StatusCode(500, new { Message = $"Lỗi khi lưu vào Database: Lỗi ràng buộc dữ liệu. Chi tiết: {ex.InnerException?.Message ?? ex.Message}" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { Message = $"Lỗi hệ thống không xác định: {ex.Message}" });
            }
        }

        [Authorize(Policy = "AddProductionDefect")]
        public async Task<IActionResult> DownloadTemplate()
        {
            var filePath = Path.Combine(_env.WebRootPath, "templates", "import_Lỗi sản xuất.xlsx");
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Không tìm thấy file mẫu.");
            }
            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = "import_Lỗi sản xuất.xlsx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        [Authorize(Policy = "ViewProductionDefect")]
        public async Task<IActionResult> ExportToExcel()
        {
            var errorDataList = await _context.ErrorDatas.ToListAsync();
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("data");
            worksheet.Cell(1, 1).Value = "STT";
            worksheet.Cell(1, 2).Value = "Order";
            worksheet.Cell(1, 3).Value = "Part Name";
            worksheet.Cell(1, 4).Value = "Qty Order";
            worksheet.Cell(1, 5).Value = "Qty NG";
            worksheet.Cell(1, 6).Value = "Ngày/Tháng";
            worksheet.Cell(1, 7).Value = "Phát hiện lỗi";
            worksheet.Cell(1, 8).Value = "Dạng lỗi";
            worksheet.Cell(1, 9).Value = "Nguyên nhân lỗi";
            worksheet.Cell(1, 10).Value = "Nội dung";
            worksheet.Cell(1, 11).Value = "Nhận định dung sai";
            worksheet.Cell(1, 12).Value = "Nguyên nhân";
            worksheet.Cell(1, 13).Value = "Đối sách";
            worksheet.Cell(1, 14).Value = "NCC";
            worksheet.Cell(1, 15).Value = "Bộ phận";
            worksheet.Cell(1, 16).Value = "NV mắc lỗi";
            worksheet.Cell(1, 17).Value = "Ngày hoàn thành giấy báo lỗi";
            worksheet.Cell(1, 18).Value = "Biện pháp khắc phục";
            worksheet.Cell(1, 19).Value = "Ghi chú";
            worksheet.Cell(1, 20).Value = "Thời gian viết giấy lỗi";
            worksheet.Cell(1, 21).Value = "Rà soát việc thực hiện NNDS";

            int row = 2;
            int stt = 1;
            foreach (var data in errorDataList)
            {
                worksheet.Cell(row, 1).Value = stt++;
                worksheet.Cell(row, 2).Value = data.OrderNo;
                worksheet.Cell(row, 3).Value = data.PartName;
                worksheet.Cell(row, 4).Value = data.QtyOrder;
                worksheet.Cell(row, 5).Value = data.QtyNG;
                worksheet.Cell(row, 6).Value = data.DateMonth.ToString("dd/MM/yyyy");
                worksheet.Cell(row, 7).Value = data.ErrorDetected;
                worksheet.Cell(row, 8).Value = data.ErrorType;
                worksheet.Cell(row, 9).Value = data.ErrorCause;
                worksheet.Cell(row, 10).Value = data.ErrorContent;
                worksheet.Cell(row, 11).Value = data.ToleranceAssessment;
                worksheet.Cell(row, 12).Value = data.Reason;
                worksheet.Cell(row, 13).Value = data.Countermeasure;
                worksheet.Cell(row, 14).Value = data.NCC;
                worksheet.Cell(row, 15).Value = data.Department;
                worksheet.Cell(row, 16).Value = data.EmployeeCode;
                worksheet.Cell(row, 17).Value = data.ErrorCompletionDate.HasValue ? data.ErrorCompletionDate.Value.ToString("dd/MM/yyyy") : string.Empty;
                worksheet.Cell(row, 18).Value = data.RemedialMeasures;
                worksheet.Cell(row, 19).Value = data.Note;
                worksheet.Cell(row, 20).Value = data.TimeWriteError;
                worksheet.Cell(row, 21).Value = data.ReviewNnds;

                row++;
            }

            worksheet.Columns().AdjustToContents();

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            var content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Thống kê lỗi.xlsx");
        }
    }
}
