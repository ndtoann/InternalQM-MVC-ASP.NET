using ExcelDataReader;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using System.Data;
using Web_QM.Models;
using Web_QM.Models.ViewModels;

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class SalaryController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public SalaryController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewSalary")]
        public async Task<IActionResult> Index(string department, string key)
        {
            var salaryQuery = _context.Salaries.AsNoTracking();
            var employeeQuery = _context.Employees.AsNoTracking();

            var resultQuery = from salary in salaryQuery
                              join employee in employeeQuery
                              on salary.EmployeeCode equals employee.EmployeeCode
                              select new
                              {
                                  Id = salary.Id,
                                  BaseSalary = salary.BaseSalary,
                                  LevelSalary = salary.LevelSalary,
                                  InsuranceSalary = salary.InsuranceSalary,
                                  EmployeeCode = employee.EmployeeCode,
                                  EmployeeName = employee.EmployeeName,
                                  Department = employee.Department
                              };

            if (!string.IsNullOrEmpty(department))
            {
                resultQuery = resultQuery.Where(item =>
                    item.Department == department
                );
            }

            if (!string.IsNullOrEmpty(key))
            {
                string searchKey = key.ToLower();
                resultQuery = resultQuery.Where(item =>
                    item.EmployeeCode.ToLower().Contains(searchKey) ||
                    item.EmployeeName.ToLower().Contains(searchKey)
                );
            }

            var res = await resultQuery.ToListAsync();

            var viewModelList = res.Select(item => new SalaryView
            {
                Id = item.Id,
                BaseSalary = item.BaseSalary,
                LevelSalary = item.LevelSalary,
                InsuranceSalary = item.InsuranceSalary,
                EmployeeCode = item.EmployeeCode,
                EmployeeName = item.EmployeeName,
                Department = item.Department
            }).ToList();

            var departments = await _context.Departments.AsNoTracking().ToListAsync();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName,
                Selected = d.DepartmentName == department
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewData["Departments"] = departmentsList;

            ViewBag.KeySearch = key;

            return View(viewModelList);
        }

        [Authorize(Policy = "ViewSalary")]
        public async Task<IActionResult> GetDetails(long id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var salaryDetails = await _context.Salaries.AsNoTracking().FirstOrDefaultAsync(s => s.Id == id);
            if (salaryDetails == null)
            {
                return NotFound();
            }
            return Json(salaryDetails);
        }

        [Authorize(Policy = "AddSalary")]
        public async Task<IActionResult> Add()
        {
            var employees = await _context.Employees.AsNoTracking().OrderBy(o => o.EmployeeCode).ToListAsync();
            var employeeList = employees.Select(e => new
            {
                Value = e.EmployeeCode,
                Text = e.EmployeeCode + " - " + e.EmployeeName
            });
            ViewBag.ListEmployee = new SelectList(employeeList, "Value", "Text");

            return View();
        }

        [Authorize(Policy = "AddSalary")]
        [HttpPost]
        public async Task<IActionResult> Add(Salary salary)
        {
            var employees = await _context.Employees.AsNoTracking().OrderBy(o => o.EmployeeCode).ToListAsync();
            var employeeList = employees.Select(e => new
            {
                Value = e.EmployeeCode,
                Text = e.EmployeeCode + " - " + e.EmployeeName
            });
            ViewBag.ListEmployee = new SelectList(employeeList, "Value", "Text");

            if (!ModelState.IsValid)
            {
                return View(salary);
            }

            var empl = await _context.Employees.AsNoTracking().FirstOrDefaultAsync(e => e.EmployeeCode == salary.EmployeeCode);
            if (empl == null)
            {
                ModelState.AddModelError("EmployeeCode", "Mã nhân viên không hợp lệ!");
                return View(salary);
            }

            bool isDuplicate = await IsDuplicate(salary.EmployeeCode);
            if (isDuplicate)
            {
                ModelState.AddModelError("EmployeeCode", "Nhân viên đã có thông tin lương!");
                return View(salary);
            }
            try
            {
                salary.InsuranceSalary = salary.InsuranceSalary ?? 0;
                salary.MealAllowance = salary.MealAllowance ?? 0;
                salary.DailyResponsibilityPay = salary.DailyResponsibilityPay ?? 0;
                salary.FuelAllowance = salary.FuelAllowance ?? 0;
                salary.ExternalWorkAllowance = salary.ExternalWorkAllowance ?? 0;
                salary.HousingSubsidy = salary.HousingSubsidy ?? 0;
                salary.DiligencePay = salary.DiligencePay ?? 0;
                salary.NoViolationBonus = salary.NoViolationBonus ?? 0;
                salary.HazardousAllowance = salary.HazardousAllowance ?? 0;
                salary.CNCStressAllowance = salary.CNCStressAllowance ?? 0;
                salary.SeniorityAllowance = salary.SeniorityAllowance ?? 0;
                salary.CertificateAllowance = salary.CertificateAllowance ?? 0;
                salary.WorkingEnvironmentAllowance = salary.WorkingEnvironmentAllowance ?? 0;
                salary.JobPositionAllowance = salary.JobPositionAllowance ?? 0;
                salary.MachineTestRunAllowance = salary.MachineTestRunAllowance ?? 0;
                salary.RVFMachineMeasurementAllowance = salary.RVFMachineMeasurementAllowance ?? 0;
                salary.StainlessSteelCleaningAllowance = salary.StainlessSteelCleaningAllowance ?? 0;
                salary.FactoryGuardAllowance = salary.FactoryGuardAllowance ?? 0;
                salary.HeavyDutyAllowance = salary.HeavyDutyAllowance ?? 0;

                salary.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Salaries.Add(salary);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(salary);
            }
        }

        [Authorize(Policy = "EditSalary")]
        public async Task<IActionResult> Edit(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var salaryToEdit = await _context.Salaries.AsNoTracking().FirstOrDefaultAsync(s => s.Id == id);
            if (salaryToEdit == null)
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

            return View(salaryToEdit);
        }

        [Authorize(Policy = "EditSalary")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, Salary salary)
        {
            if(id != salary.Id)
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
            if (!ModelState.IsValid)
            {
                return View(salary);
            }
            var oldSalary = await _context.Salaries.AsNoTracking().FirstOrDefaultAsync(s => s.Id == salary.Id);
            if (oldSalary == null)
            {
                return NotFound();
            }
            var empl = await _context.Employees.AsNoTracking().FirstOrDefaultAsync(e => e.EmployeeCode == salary.EmployeeCode);
            if (empl == null)
            {
                ModelState.AddModelError("EmployeeCode", "Mã nhân viên không hợp lệ!");
                return View(salary);
            }
            bool isDuplicate = await IsDuplicate(salary.EmployeeCode, salary.Id);
            if (isDuplicate)
            {
                ModelState.AddModelError("EmployeeCode", "Nhân viên đã có thông tin lương!");
                return View(salary);
            }
            try
            {
                salary.InsuranceSalary = salary.InsuranceSalary ?? 0;
                salary.MealAllowance = salary.MealAllowance ?? 0;
                salary.DailyResponsibilityPay = salary.DailyResponsibilityPay ?? 0;
                salary.FuelAllowance = salary.FuelAllowance ?? 0;
                salary.ExternalWorkAllowance = salary.ExternalWorkAllowance ?? 0;
                salary.HousingSubsidy = salary.HousingSubsidy ?? 0;
                salary.DiligencePay = salary.DiligencePay ?? 0;
                salary.NoViolationBonus = salary.NoViolationBonus ?? 0;
                salary.HazardousAllowance = salary.HazardousAllowance ?? 0;
                salary.CNCStressAllowance = salary.CNCStressAllowance ?? 0;
                salary.SeniorityAllowance = salary.SeniorityAllowance ?? 0;
                salary.CertificateAllowance = salary.CertificateAllowance ?? 0;
                salary.WorkingEnvironmentAllowance = salary.WorkingEnvironmentAllowance ?? 0;
                salary.JobPositionAllowance = salary.JobPositionAllowance ?? 0;
                salary.MachineTestRunAllowance = salary.MachineTestRunAllowance ?? 0;
                salary.RVFMachineMeasurementAllowance = salary.RVFMachineMeasurementAllowance ?? 0;
                salary.StainlessSteelCleaningAllowance = salary.StainlessSteelCleaningAllowance ?? 0;
                salary.FactoryGuardAllowance = salary.FactoryGuardAllowance ?? 0;
                salary.HeavyDutyAllowance = salary.HeavyDutyAllowance ?? 0;
                salary.CreatedDate = oldSalary.CreatedDate;
                salary.UpdatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Salaries.Update(salary);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(salary);
            }
        }

        [Authorize(Policy = "DeleteSalary")]
        public async Task<IActionResult> Delete(long id)
        {
            if(id == null)
            {
                return NotFound();
            }

            var salaryToDelete = await _context.Salaries.FindAsync(id);
            if (salaryToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy dữ liệu!";
                return RedirectToAction(nameof(Index));
            }
            try
            {
                _context.Salaries.Remove(salaryToDelete);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa dữ liệu!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Index));
        }

        [Authorize(Policy = "AddSalary")]
        [HttpPost]
        public async Task<IActionResult> Import(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest(new { Message = "Vui lòng chọn một file Excel." });
            }

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var salaryToInsert = new List<Salary>();

            var requiredFields = new[] { "Mã NV", "Bậc lương", "Lương CB" };

            var notFoundEmployeeCount = 0;
            var totalRowsProcessed = 0;
            var insertedCount = 0;
            var skippedCount = 0;

            var employeeCodeInFileHashSet = new HashSet<string>();
            var duplicateInFileCount = 0;

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
                        totalRowsProcessed = table.Rows.Count;

                        var missingColumns = requiredFields.Where(field => !table.Columns.Contains(field)).ToList();
                        if (missingColumns.Any())
                        {
                            return BadRequest(new { Message = $"File Excel thiếu các cột bắt buộc: {string.Join(", ", missingColumns)}" });
                        }

                        var employeeCodesInFile = table.AsEnumerable()
                            .Select(row => row["Mã NV"]?.ToString()?.Trim())
                            .Where(code => !string.IsNullOrWhiteSpace(code))
                            .Distinct()
                            .ToList();

                        var existingEmployees = await _context.Employees
                            .Where(e => employeeCodesInFile.Contains(e.EmployeeCode))
                            .Select(e => e.EmployeeCode)
                            .ToListAsync();

                        var existingSalaryEmployeeCodes = await _context.Salaries
                            .Where(s => employeeCodesInFile.Contains(s.EmployeeCode))
                            .Select(s => s.EmployeeCode)
                            .ToListAsync();

                        var existingSalaryHashSet = new HashSet<string>(existingSalaryEmployeeCodes);

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
                                return BadRequest(new { Message = $"Lỗi tại dòng {rowIndex}: Các trường bắt buộc bị thiếu dữ liệu: {string.Join(", ", missingDataFields)}" });
                            }

                            try
                            {
                                var employeeCode = row["Mã NV"].ToString().Trim();

                                if (employeeCodeInFileHashSet.Contains(employeeCode))
                                {
                                    duplicateInFileCount++;
                                    continue;
                                }
                                employeeCodeInFileHashSet.Add(employeeCode);

                                if (!existingEmployees.Contains(employeeCode))
                                {
                                    notFoundEmployeeCount++;
                                    continue;
                                }

                                // KIỂM TRA TRÙNG LẶP DB VÀ BỎ QUA
                                if (existingSalaryHashSet.Contains(employeeCode))
                                {
                                    skippedCount++;
                                    continue;
                                }

                                var newSalaryData = MapSalaryFromExcelRow(row);
                                newSalaryData.EmployeeCode = employeeCode;
                                newSalaryData.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");

                                salaryToInsert.Add(newSalaryData);
                                insertedCount++;
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
                if (salaryToInsert.Any())
                {
                    await _context.Salaries.AddRangeAsync(salaryToInsert);
                    await _context.SaveChangesAsync();
                }

                var message = $"Hoàn tất import. Tổng số dòng trong file: {totalRowsProcessed}. Đã thêm mới thành công: {insertedCount} bản ghi. Đã bỏ qua do trùng Mã NV trong file: {duplicateInFileCount} bản ghi. Đã bỏ qua do Mã nhân viên không tồn tại trong hệ thống: {notFoundEmployeeCount} bản ghi. Đã bỏ qua do trùng Mã NV trong hệ thống: {skippedCount} bản ghi.";

                return Ok(new { Message = message, InsertedCount = insertedCount, DuplicateInFileCount = duplicateInFileCount, NotFoundEmployeeCount = notFoundEmployeeCount, SkippedCount = skippedCount });
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

        private Salary MapSalaryFromExcelRow(DataRow row)
        {
            return new Salary
            {
                LevelSalary = Convert.ToInt32(row["Bậc lương"]),
                BaseSalary = Convert.ToDecimal(row["Lương CB"]),
                InsuranceSalary = GetDecimalValue(row, "Lương đóng BH"),
                MealAllowance = GetDecimalValue(row, "Tiền ăn"),
                DailyResponsibilityPay = GetDecimalValue(row, "Tiền trách nhiệm ngày"),
                FuelAllowance = GetDecimalValue(row, "Phụ cấp xăng xe"),
                ExternalWorkAllowance = GetDecimalValue(row, "Phụ cấp làm ngoài"),
                HousingSubsidy = GetDecimalValue(row, "Trợ cấp nhà trọ"),
                DiligencePay = GetDecimalValue(row, "Tiền chuyên cần"),
                NoViolationBonus = GetDecimalValue(row, "Không phạm lỗi"),
                HazardousAllowance = GetDecimalValue(row, "Phụ cấp độc hại"),
                CNCStressAllowance = GetDecimalValue(row, "Phụ cấp căng thẳng CNC"),
                SeniorityAllowance = GetDecimalValue(row, "Phụ cấp thâm niên"),
                CertificateAllowance = GetDecimalValue(row, "Phụ cấp bằng"),
                WorkingEnvironmentAllowance = GetDecimalValue(row, "Phụ cấp môi trường làm việc"),
                JobPositionAllowance = GetDecimalValue(row, "Phụ cấp vị trí làm việc"),
                MachineTestRunAllowance = GetDecimalValue(row, "Phụ cấp chạy thử máy"),
                RVFMachineMeasurementAllowance = GetDecimalValue(row, "Phụ cấp đo máy RVF"),
                StainlessSteelCleaningAllowance = GetDecimalValue(row, "Phụ cấp rửa hàng Inox"),
                FactoryGuardAllowance = GetDecimalValue(row, "Phụ cấp trông xưởng"),
                HeavyDutyAllowance = GetDecimalValue(row, "Phụ cấp nặng nhọc")
            };
        }

        private decimal? GetDecimalValue(DataRow row, string columnName)
        {
            if (!row.Table.Columns.Contains(columnName) || row[columnName] == DBNull.Value || string.IsNullOrWhiteSpace(row[columnName].ToString()))
            {
                return 0;
            }

            if (decimal.TryParse(row[columnName].ToString(), out decimal result))
            {
                return result;
            }
            return 0;
        }

        [Authorize(Policy = "AddSalary")]
        public async Task<IActionResult> DownloadTemplate()
        {
            var filePath = Path.Combine(_env.WebRootPath, "templates", "import_lương.xlsx");
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Không tìm thấy file mẫu.");
            }
            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = "import_lương.xlsx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        private async Task<bool> IsDuplicate(string employeeCode, long currentId = 0)
        {
            bool exists = await _context.Salaries
                .AnyAsync(p =>
                    p.EmployeeCode == employeeCode &&
                    p.Id != currentId
                );
            return exists;
        }
    }
}
