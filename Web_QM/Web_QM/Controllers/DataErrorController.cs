using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using ExcelDataReader;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Data;
using System.Globalization;
using System.Threading.Tasks;
using Web_QM.Models;

namespace Web_QM.Controllers
{
    public class DataErrorController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public DataErrorController( QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewProductionDefect")]
        public async Task<IActionResult> Index()
        {
            return View();
        }

        [Authorize(Policy = "ViewProductionDefect")]
        public async Task<IActionResult> GetDataErrors()
        {
            var errors = await _context.ErrorDatas.AsNoTracking().OrderByDescending(o => o.DateMonth).ToListAsync();
            return Json(errors);
        }

        [Authorize(Policy = "AddProductionDefect")]
        [HttpPost]
        public async Task<IActionResult> Add([FromBody] ErrorData errorData)
        {
            if (!ModelState.IsValid)
            {
                return Json(new { success = false, message = "Vui lòng kiểm tra lại dữ liệu." });
            }

            try
            {
                _context.ErrorDatas.Add(errorData);
                await _context.SaveChangesAsync();
                return Json(new { success = true, message = "Đã lưu dữ liệu." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }

        [Authorize(Policy = "EditProductionDefect")]
        [HttpPost]
        public async Task<IActionResult> Update([FromBody] ErrorData errorData)
        {
            if (!ModelState.IsValid)
            {
                return Json(new { success = false, message = "Vui lòng kiểm tra lại dữ liệu." });
            }

            try
            {
                _context.ErrorDatas.Update(errorData);
                await _context.SaveChangesAsync();
                return Json(new { success = true, message = "Đã cập nhật dữ liệu." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }

        [Authorize(Policy = "DeleteProductionDefect")]
        public async Task<IActionResult> Delete(long id)
        {
            if(id == null)
            {
                return Json(new { success = false, message = "Thông tin xóa không tồn tại." });
            }
            var errorToDelete = await _context.ErrorDatas.FirstOrDefaultAsync(e => e.Id ==id);
            if (errorToDelete == null)
            {
                return Json(new { success = false, message = "Thông tin xóa không tồn tại." });
            }
            try
            {
                _context.ErrorDatas.Remove(errorToDelete);
                await _context.SaveChangesAsync();
                return Json(new { success = true, message = "Đã xóa dữ liệu." });
            }
            catch(Exception ex)
            {
                return Json(new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }

        [Authorize(Policy = "AddProductionDefect")]
        [HttpPost]
        public async Task<IActionResult> ImportExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("Vui lòng chọn một file Excel.");
            }

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var errorDataList = new List<ErrorData>();

            try
            {
                using (var stream = file.OpenReadStream())
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

                        if (dataSet.Tables.Count == 0)
                        {
                            return BadRequest("File Excel không có dữ liệu.");
                        }

                        DataTable table = dataSet.Tables[0];
                        var requiredFields = new[] { "OrderNo", "PartName", "QtyOrder", "QtyNG", "Ngày Tháng", "Phát hiện lỗi", "Dạng lỗi", "Nội dung", "NCC" };

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            var row = table.Rows[i];
                            var missingFields = new List<string>();
                            foreach (var field in requiredFields)
                            {
                                if (row[field] == null || string.IsNullOrWhiteSpace(row[field].ToString()))
                                {
                                    missingFields.Add(field);
                                }
                            }

                            if (missingFields.Any())
                            {
                                return BadRequest(new { Message = $"Lỗi tại dòng {i + 2}: Các trường bắt buộc bị thiếu dữ liệu: {string.Join(", ", missingFields)}." });
                            }

                            try
                            {
                                var errorData = new ErrorData
                                {
                                    OrderNo = row["OrderNo"].ToString(),
                                    PartName = row["PartName"].ToString(),
                                    QtyOrder = int.Parse(row["QtyOrder"].ToString()),
                                    QtyNG = int.Parse(row["QtyNG"].ToString()),

                                    DateMonth = DateOnly.FromDateTime(DateTime.Parse(row["Ngày Tháng"].ToString())),
                                    ErrorDetected = row["Phát hiện lỗi"].ToString(),
                                    ErrorType = row["Dạng lỗi"].ToString(),
                                    ErrorCause = row["Nguyên nhân lỗi"]?.ToString(),
                                    ErrorContent = row["Nội dung"].ToString(),
                                    ToleranceAssessment = row["Nhận định dung sai"]?.ToString(),
                                    Reason = row["Nguyên nhân"]?.ToString(),
                                    Countermeasure = row["Đối sách"]?.ToString(),
                                    NCC = row["NCC"].ToString(),
                                    EmployeeCode = row["Mã nhân viên"]?.ToString(),
                                    Department = row["Bộ phận"]?.ToString(),
                                    ErrorCompletionDate = GetDateOnly(row["Ngày hoàn thành giấy báo lỗi"]),
                                    RemedialMeasures = row["Biện pháp khắc phục"]?.ToString(),
                                    Note = row["Ghi chú"]?.ToString(),
                                    TimeWriteError = row["Thời gian viết giấy lỗi"]?.ToString(),
                                    ReviewNnds = row["Rà soát việc thực hiện NNDS"]?.ToString()
                                };
                                errorDataList.Add(errorData);
                            }
                            catch (Exception ex)
                            {
                                return BadRequest(new { Message = $"Lỗi tại dòng {i + 2}: Lỗi chuyển đổi dữ liệu - {ex.Message}" });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Có lỗi xảy ra khi đọc file: {ex.Message}");
            }
            await _context.ErrorDatas.AddRangeAsync(errorDataList);
            await _context.SaveChangesAsync();

            return Ok("Lưu dữ liệu thành công.");
        }

        private DateOnly? GetDateOnly(object value)
        {
            if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
            {
                return new DateOnly(1, 1, 1);
            }

            if (DateTime.TryParse(value.ToString(), out var dateTime))
            {
                return DateOnly.FromDateTime(dateTime);
            }

            return new DateOnly(1, 1, 1);
        }

        [Authorize(Policy = "AddProductionDefect")]
        public async Task<IActionResult> DownloadTemplate()
        {
            var filePath = Path.Combine(_env.WebRootPath, "templates", "import_lỗi sx.xlsx");
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Không tìm thấy file mẫu.");
            }
            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = "import_lỗi sx.xlsx";
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
                worksheet.Cell(row, 17).Value = data.ErrorCompletionDate.HasValue && data.ErrorCompletionDate.Value != new DateOnly(1, 1, 1) ? data.ErrorCompletionDate.Value.ToString("dd/MM/yyyy") : string.Empty;
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

        [Authorize(Policy = "ViewProductionDefect")]
        public async Task<IActionResult> GetErrorChartData(int year)
        {
            int currentYear = year;

            var errors = await _context.ErrorDatas
                                       .Where(e => e.DateMonth.Year == currentYear)
                                       .ToListAsync();

            var monthlyData = errors.GroupBy(e => e.DateMonth.Month)
                                    .OrderBy(g => g.Key)
                                    .ToDictionary(g => g.Key, g => g.ToList());

            var chartData = new List<object>();

            for (int month = 1; month <= 12; month++)
            {
                int qmErrorsCount = 0;
                int otherNccErrorsCount = 0;

                if (monthlyData.ContainsKey(month))
                {
                    var monthErrors = monthlyData[month];

                    qmErrorsCount = monthErrors.Count(e => e.NCC.Trim().ToUpper() == "QM");
                    otherNccErrorsCount = monthErrors.Count(e => e.NCC.Trim().ToUpper() != "QM");
                }

                var monthData = new
                {
                    Month = $"T{month}",
                    QMErrors = qmErrorsCount,
                    OtherNccErrors = otherNccErrorsCount
                };
                chartData.Add(monthData);
            }

            return Ok(chartData);
        }

        [Authorize(Policy = "ViewProductionDefect")]
        public async Task<IActionResult> GetLineChartData(int year)
        {
            var currentYear = year;

            var data = await _context.ErrorDatas
                .Where(e => e.DateMonth.Year == currentYear)
                .GroupBy(e => new { Ncc = e.NCC, Month = e.DateMonth.Month })
                .Select(g => new
                {
                    NCC = g.Key.Ncc,
                    Month = g.Key.Month,
                    TotalErrors = g.Count()
                })
                .ToListAsync();

            var nccs = data.Select(d => d.NCC).Distinct().ToList();

            var datasets = nccs.Select(ncc => new
            {
                label = ncc,
                data = Enumerable.Range(1, 12)
                    .Select(month => data.FirstOrDefault(d => d.NCC == ncc && d.Month == month)?.TotalErrors ?? 0)
                    .ToList()
            }).ToList();

            var chartData = new
            {
                labels = Enumerable.Range(1, 12).Select(m => $"Tháng {m}").ToList(),
                datasets = datasets
            };

            return Ok(chartData);
        }

        [Authorize(Policy = "ClientViewViolation5S")]
        public async Task<IActionResult> Violation5S()
        {
            ViewBag.Departments = await _context.Departments
                                                .Select(d => new { d.DepartmentName })
                                                .ToListAsync();
            return View();
        }

        [Authorize(Policy = "ClientViewViolation5S")]
        public async Task<IActionResult> GetViolations(string key, string department, DateTime? startDate = null, DateTime? endDate = null)
        {
            try
            {
                string searchKey = (key ?? "").ToLower();
                string filterDepartment = department ?? "";
                DateOnly? startOnly = startDate.HasValue ? DateOnly.FromDateTime(startDate.Value.Date) : (DateOnly?)null;
                DateOnly? endOnly = endDate.HasValue ? DateOnly.FromDateTime(endDate.Value.Date) : (DateOnly?)null;

                var query = _context.EmployeeViolation5S
                    .AsNoTracking()
                    .Join(_context.Employees,
                          v => v.EmployeeCode,
                          e => e.EmployeeCode,
                          (v, e) => new { Violation = v, Employee = e })
                    .Join(_context.Violation5S,
                          ve => ve.Violation.Violation5SId,
                          vs => vs.Id,
                          (ve, vs) => new { ve.Violation, ve.Employee, ViolationContent = vs });

                query = query.Where(x =>
                    (string.IsNullOrEmpty(searchKey) ||
                     (x.Employee.EmployeeCode != null && x.Employee.EmployeeCode.ToLower().Contains(searchKey)) ||
                     (x.Employee.EmployeeName != null && x.Employee.EmployeeName.ToLower().Contains(searchKey))
                    )
                    &&
                    (string.IsNullOrEmpty(filterDepartment) || x.Employee.Department == filterDepartment)
                    &&
                    (!startOnly.HasValue || x.Violation.DateMonth >= startOnly.Value)
                    &&
                    (!endOnly.HasValue || x.Violation.DateMonth <= endOnly.Value)
                );

                var violations = await query
                    .OrderByDescending(x => x.Violation.DateMonth)
                    .Take(1000)
                    .Select(x => new
                    {
                        Id = x.Violation.Id,
                        EmployeeCode = x.Violation.EmployeeCode,
                        EmployeeName = x.Employee.EmployeeName,
                        Department = x.Employee.Department,
                        Content5S = x.ViolationContent.Content5S,
                        DateMonth = x.Violation.DateMonth,
                        Qty = x.Violation.Qty
                    })
                    .ToListAsync();

                return Json(new { success = true, data = violations });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }

        [Authorize(Policy = "ClientViewViolation5S")]
        public async Task<IActionResult> GetChartData(int year)
        {
            try
            {
                var monthlyData = await _context.EmployeeViolation5S
                    .Where(v => v.DateMonth.Year == year)
                    .GroupBy(v => v.DateMonth.Month)
                    .Select(g => new 
                    {
                        Month = g.Key,
                        TotalQty = g.Sum(v => v.Qty)
                    })
                    .OrderBy(m => m.Month)
                    .ToListAsync();

                var departmentData = await _context.EmployeeViolation5S
                    .Where(v => v.DateMonth.Year == year)
                    .Join(_context.Employees,
                          v => v.EmployeeCode,
                          e => e.EmployeeCode,
                          (v, e) => new { Violation = v, Employee = e })
                    .Where(x => x.Employee.Department != null)
                    .GroupBy(x => x.Employee.Department)
                    .Select(g => new
                    {
                        Department = g.Key,
                        TotalQty = g.Sum(x => x.Violation.Qty)
                    })
                    .OrderByDescending(d => d.TotalQty)
                    .ToListAsync();

                var chartData = new
                {
                    MonthlyData = monthlyData,
                    DepartmentData = departmentData
                };

                return Json(new { success = true, data = chartData });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = $"Lỗi khi tải dữ liệu biểu đồ: {ex.Message}" });
            }
        }
    }
}
