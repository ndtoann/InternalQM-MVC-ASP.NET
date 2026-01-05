using System.ComponentModel.DataAnnotations;
using System.Data;
using ExcelDataReader;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;

namespace Web_QM.Areas.HR.Controllers
{
    [Area("HR")]
    public class KaizenController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public KaizenController(QMContext context, IWebHostEnvironment env  )
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewKaizen")]
        public async Task<IActionResult> Index(string key, string? review, DateTime? startDate = null, DateTime? endDate = null)
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
            var res = await _context.Kaizens.AsNoTracking()
                                            .Where(k =>
                                                (string.IsNullOrEmpty(key) ||
                                                (k.EmployeeCode.ToLower().Contains(key.ToLower())) ||
                                                (k.EmployeeName.ToLower().Contains(key.ToLower())) ||
                                                (k.ProposedIdea.ToLower().Contains(key.ToLower()))) &&
                                                (string.IsNullOrEmpty(review) ||
                                                (k.ManagementReview != null && k.ManagementReview.ToLower().Contains(review.ToLower())) &&
                                                (k.DateMonth >= DateOnly.FromDateTime(startDate.Value)) &&
                                                (k.DateMonth <= DateOnly.FromDateTime(endDate.Value)))
                                            )
                                            .OrderByDescending(o => o.DateMonth)
                                            .Take(1000)
                                            .ToListAsync();
            var statusOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "", Text = "Đánh giá BQL" },
                new SelectListItem { Value = "A", Text = "A" },
                new SelectListItem { Value = "B", Text = "B" },
                new SelectListItem { Value = "C", Text = "C" },
                new SelectListItem { Value = "D", Text = "D" }
            };

            ViewData["ReviewList"] = new SelectList(statusOptions, "Value", "Text", review);
            ViewBag.KeySearch = key;
            ViewBag.StartDate = startDate.Value.ToString("yyyy-MM-dd");
            ViewBag.EndDate = endDate.Value.Date.ToString("yyyy-MM-dd");
            return View(res);
        }

        [Authorize(Policy = "AddKaizen")]
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
            ViewData["Departments"] = departmentsList;

            return View();
        }

        [Authorize(Policy = "AddKaizen")]
        [HttpPost]
        public async Task<IActionResult> Add(Kaizen kaizem, IFormFile kaizenFile)
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
            ViewData["Departments"] = departmentsList;

            var empl = await _context.Employees.AsNoTracking().FirstOrDefaultAsync(e => e.EmployeeCode == kaizem.EmployeeCode);
            if (empl == null)
            {
                TempData["ErrorMessage"] = "Mã nhân viên không hợp lệ";
                return View(kaizem);
            }
            try
            {
                kaizem.EmployeeName = empl.EmployeeName;
                kaizem.Department = empl.Department;
                string kaizanePic = UploadImg(kaizenFile);
                kaizem.Picture = kaizanePic;
                _context.Kaizens.Add(kaizem);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(kaizem);
            }
        }

        [Authorize(Policy = "EditKaizen")]
        public async Task<IActionResult> Edit(long id)
        {
            if(id == null)
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
            ViewData["Departments"] = departmentsList;

            var kaizenToEdit = await _context.Kaizens.AsNoTracking().FirstOrDefaultAsync(k => k.Id == id);
            if (kaizenToEdit == null)
            {
                return NotFound();
            }
            return View(kaizenToEdit);
        }

        [Authorize(Policy = "EditKaizen")]
        [HttpPost]
        public async Task<IActionResult> Edit (long id, Kaizen kaizem, IFormFile kaizenFile)
        {
            if (id != kaizem.Id)
            {
                return NotFound();
            }
            var kaizenToEdit = await _context.Kaizens.AsNoTracking().FirstOrDefaultAsync(k => k.Id == kaizem.Id);
            if (kaizenToEdit == null)
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
            ViewData["Departments"] = departmentsList;

            var empl = await _context.Employees.AsNoTracking().FirstOrDefaultAsync(e => e.EmployeeCode == kaizem.EmployeeCode);
            if (empl == null)
            {
                TempData["ErrorMessage"] = "Mã nhân viên không hợp lệ";
                return View(kaizem);
            }
            try
            {
                if (kaizenFile != null && kaizenFile.Length > 0)
                {
                    if (!string.IsNullOrEmpty(kaizenToEdit.Picture))
                    {
                        string oldFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "imgs", "kaizwns", kaizenToEdit.Picture);
                        if (System.IO.File.Exists(oldFilePath))
                        {
                            System.IO.File.Delete(oldFilePath);
                        }
                    }

                    string newAvatarName = UploadImg(kaizenFile);
                    kaizem.Picture = newAvatarName;
                }
                else
                {
                    kaizem.Picture = kaizenToEdit.Picture;
                }
                kaizem.EmployeeName = empl.EmployeeName;
                kaizem.Department = empl.Department;
                _context.Kaizens.Update(kaizem);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(kaizem);
            }
        }

        [Authorize(Policy = "DeleteKaizen")]
        public async Task<IActionResult> Delete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var kaizenToDelete = await _context.Kaizens.FindAsync(id);

            if (kaizenToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy kaizen!";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                if (kaizenToDelete.Picture != null)
                {
                    string oldFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "imgs", "kaizens", kaizenToDelete.Picture);
                    if (System.IO.File.Exists(oldFilePath))
                    {
                        System.IO.File.Delete(oldFilePath);
                    }
                }
                _context.Kaizens.Remove(kaizenToDelete);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa kaizen!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Index));
        }

        private string UploadImg(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return null;
            }

            var allowedExtensions = new[] { ".png", ".jpeg", ".jpg" };
            const int maxFileSizeMB = 2;
            const long maxFileSizeInBytes = maxFileSizeMB * 1024 * 1024;
            if (file.Length > maxFileSizeInBytes)
            {
                Console.WriteLine($"Lỗi: Dung lượng file quá lớn. Tối đa là {maxFileSizeMB}MB.");
                return null;
            }

            try
            {
                string fileExtension = Path.GetExtension(file.FileName).ToLowerInvariant();

                if (!allowedExtensions.Contains(fileExtension))
                {
                    Console.WriteLine($"Lỗi: Định dạng file không được hỗ trợ. Chỉ cho phép các định dạng: {string.Join(", ", allowedExtensions)}.");
                    return null;
                }

                string newFileName = $"kaizen_{Guid.NewGuid().ToString()}{fileExtension}";
                string uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "imgs", "kaizens");
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }
                string filePath = Path.Combine(uploadPath, newFileName);
                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(fileStream);
                }

                return newFileName;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi tải ảnh lên: {ex.Message}");
                return null;
            }
        }

        [Authorize(Policy = "AddKaizen")]
        [HttpPost]
        public async Task<IActionResult> Import(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest(new { Message = "Vui lòng chọn một file Excel." });
            }

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var kaizenList = new List<Kaizen>();
            var requiredFields = new[] { "DateMonth", "EmployeeCode", "ImprovementTitle", "ProposedIdea", "EmployeeName", "Department" };

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
                                object dateValue = row["DateMonth"];

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

                                var kaizen = new Kaizen
                                {
                                    Id = 0,
                                    DateMonth = dateMonth,
                                    EmployeeCode = row["EmployeeCode"].ToString().Trim(),

                                    EmployeeName = row["EmployeeName"].ToString().Trim(),
                                    Department = row["Department"].ToString().Trim(),
                                    AppliedDepartment = row.Table.Columns.Contains("AppliedDepartment") && row["AppliedDepartment"] != DBNull.Value ? row["AppliedDepartment"].ToString().Trim() : null,
                                    ImprovementGoal = row.Table.Columns.Contains("ImprovementGoal") && row["ImprovementGoal"] != DBNull.Value ? row["ImprovementGoal"].ToString().Trim() : null,
                                    ImprovementTitle = row["ImprovementTitle"].ToString().Trim(),
                                    CurrentSituation = row.Table.Columns.Contains("CurrentSituation") && row["CurrentSituation"] != DBNull.Value ? row["CurrentSituation"].ToString().Trim() : null,
                                    ProposedIdea = row["ProposedIdea"].ToString().Trim(),
                                    EstimatedBenefit = row.Table.Columns.Contains("EstimatedBenefit") && row["EstimatedBenefit"] != DBNull.Value ? row["EstimatedBenefit"].ToString().Trim() : null,
                                    TeamLeaderRating = row.Table.Columns.Contains("TeamLeaderRating") && row["TeamLeaderRating"] != DBNull.Value ? row["TeamLeaderRating"].ToString().Trim() : null,
                                    ManagementReview = row.Table.Columns.Contains("ManagementReview") && row["ManagementReview"] != DBNull.Value ? row["ManagementReview"].ToString().Trim() : null,
                                    Picture = row.Table.Columns.Contains("Picture") && row["Picture"] != DBNull.Value ? row["Picture"].ToString().Trim() : null,
                                    Deadline = row.Table.Columns.Contains("Deadline") && row["Deadline"] != DBNull.Value ? row["Deadline"].ToString().Trim() : null,
                                    StartTime = row.Table.Columns.Contains("StartTime") && row["StartTime"] != DBNull.Value ? row["StartTime"].ToString().Trim() : null,
                                    CurrentStatus = row.Table.Columns.Contains("CurrentStatus") && row["CurrentStatus"] != DBNull.Value ? row["CurrentStatus"].ToString().Trim() : null,
                                    Note = row.Table.Columns.Contains("Note") && row["Note"] != DBNull.Value ? row["Note"].ToString().Trim() : null,
                                };

                                var validationContext = new ValidationContext(kaizen, serviceProvider: null, items: null);
                                var validationResults = new List<ValidationResult>();

                                if (!Validator.TryValidateObject(kaizen, validationContext, validationResults, true))
                                {
                                    var errors = string.Join("; ", validationResults.Select(r => r.ErrorMessage));
                                    return BadRequest(new { Message = $"Lỗi tại dòng {rowIndex}: Lỗi Validation Model. Chi tiết: {errors}" });
                                }

                                kaizenList.Add(kaizen);
                            }
                            catch (Exception ex)
                            {
                                return BadRequest(new { Message = $"Lỗi tại dòng {rowIndex}: Lỗi chuyển đổi dữ liệu. Chi tiết: {ex.Message}" });
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
                if (kaizenList.Any())
                {
                    await _context.Kaizens.AddRangeAsync(kaizenList);
                    await _context.SaveChangesAsync();
                }

                return Ok(new { Message = $"Đã import và lưu thành công {kaizenList.Count} bản ghi Kaizen." });
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

        [Authorize(Policy = "AddKaizen")]
        public async Task<IActionResult> DownloadTemplate()
        {
            var filePath = Path.Combine(_env.WebRootPath, "templates", "import_Kaizen.xlsx");
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Không tìm thấy file mẫu.");
            }
            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = "import_Kaizen.xlsx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }
}
