using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;
using Web_QM.Models.ViewModels;

namespace Web_QM.Areas.HR.Controllers
{
    [Area("HR")]
    public class EmployeeController : Controller
    {
        private readonly QMContext _context;

        public EmployeeController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ViewEmployee")]
        public async Task<IActionResult> Index(string department, string key_search)
        {
            var res = await _context.Employees
                           .AsNoTracking()
                           .Where(e => (string.IsNullOrEmpty(department) || e.Department == department) &&
                                       (string.IsNullOrEmpty(key_search) || (e.EmployeeCode != null && e.EmployeeCode.ToLower().Contains(key_search.ToLower())) ||
                                        (e.EmployeeName != null && e.EmployeeName.ToLower().Contains(key_search.ToLower()))))
                           .OrderBy(e => e.EmployeeCode)
                           .ToListAsync();

            var departments = await _context.Departments.AsNoTracking().ToListAsync();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName,
                Selected = d.DepartmentName == department
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận/phòng ban" });
            ViewData["Departments"] = departmentsList;
            ViewBag.KeySearch = key_search;
            return View(res);
        }

        [Authorize(Policy = "AddEmployee")]
        public async Task<IActionResult> Add()
        {
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

        [Authorize(Policy = "AddEmployee")]
        [HttpPost]
        public async Task<IActionResult> Add(Employee employee, IFormFile avatarFile)
        {
            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewData["Departments"] = departmentsList;

            var existingEmployee = await _context.Employees.AnyAsync(e => e.EmployeeCode == employee.EmployeeCode);
            if (existingEmployee)
            {
                TempData["ErrorMessage"] = "Mã nhân viên đã tồn tại!";
                return View(employee);
            }

            try
            {
                string avatarName = UploadAvatar(avatarFile, employee.EmployeeCode);
                employee.Avatar = avatarName;

                employee.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");

                _context.Employees.Add(employee);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";

                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!";
                return View(employee);
            }
        }

        [Authorize(Policy = "EditEmployee")]
        public async Task<IActionResult> Edit(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewData["Departments"] = departmentsList;

            var employeeToEdit = await _context.Employees.FirstOrDefaultAsync(x => x.Id == id);
            if (employeeToEdit == null)
            {
                return NotFound();
            }

            return View(employeeToEdit);
        }

        [Authorize(Policy = "EditEmployee")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, Employee employee, IFormFile avatarFile)
        {
            if (id != employee.Id)
            {
                return NotFound();
            }
            var emplToEdit = await _context.Employees.AsNoTracking().FirstOrDefaultAsync(e => e.Id == employee.Id);
            if (emplToEdit == null)
            {
                return NotFound();
            }

            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewData["Departments"] = departmentsList;

            var existingEmployee = await _context.Employees.AnyAsync(e => e.EmployeeCode == employee.EmployeeCode && e.Id != employee.Id);
            if (existingEmployee)
            {
                TempData["ErrorMessage"] = "Mã nhân viên đã tồn tại!";
                return View(employee);
            }
            try
            {
                if (avatarFile != null && avatarFile.Length > 0)
                {
                    if (!string.IsNullOrEmpty(emplToEdit.Avatar))
                    {
                        string oldFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "imgs", "avatars", emplToEdit.Avatar);
                        if (System.IO.File.Exists(oldFilePath))
                        {
                            System.IO.File.Delete(oldFilePath);
                        }
                    }

                    string newAvatarName = UploadAvatar(avatarFile, employee.EmployeeCode);
                    employee.Avatar = newAvatarName;
                }
                else
                {
                    employee.Avatar = emplToEdit.Avatar;
                }
                employee.UpdatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Employees.Update(employee);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!";
                return View(employee);
            }
        }

        private string UploadAvatar(IFormFile file, string emplCode)
        {
            if (file == null || file.Length == 0)
            {
                return null;
            }

            var allowedExtensions = new[] { ".png", ".jpeg", ".jpg" };
            const int maxFileSizeMB = 10;
            const long maxFileSizeInBytes = maxFileSizeMB * 1024 * 1024;
            if (file.Length > maxFileSizeInBytes)
            {
                return null;
            }

            try
            {
                string fileExtension = Path.GetExtension(file.FileName).ToLowerInvariant();

                if (!allowedExtensions.Contains(fileExtension))
                {
                    return null;
                }

                string newFileName = $"avatar_{emplCode}{fileExtension}";
                string uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "imgs", "avatars");
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
                return null;
            }
        }

        [Authorize(Policy = "DeleteEmployee")]
        public async Task<IActionResult> Delete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var employeeToDelete = await _context.Employees.FirstOrDefaultAsync(x => x.Id == id);

            if (employeeToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy nhân viên này!";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                var acc = await _context.Accounts.FirstOrDefaultAsync(a => a.StaffCode == employeeToDelete.EmployeeCode);
                if(acc != null)
                {
                    await _context.AccountPermissions
                     .Where(a => a.AccountId == acc.Id)
                     .ExecuteDeleteAsync();
                }
                await _context.Accounts
                     .Where(a => a.StaffCode == employeeToDelete.EmployeeCode)
                     .ExecuteDeleteAsync();

                await _context.ExamPeriodicAnswers
                     .Where(a => a.EmployeeCode == employeeToDelete.EmployeeCode)
                     .ExecuteDeleteAsync();

                await _context.ExamTrialRunAnswers
                     .Where(a => a.EmployeeCode == employeeToDelete.EmployeeCode)
                     .ExecuteDeleteAsync();

                await _context.ExamTrainingAnswers
                     .Where(a => a.EmployeeCode == employeeToDelete.EmployeeCode)
                     .ExecuteDeleteAsync();

                await _context.EmployeeTrainingResults
                     .Where(a => a.EmployeeId == employeeToDelete.Id)
                     .ExecuteDeleteAsync();

                await _context.Feedbacks
                     .Where(a => a.EmployeeId == employeeToDelete.Id)
                     .ExecuteDeleteAsync();

                await _context.TestPractices
                     .Where(a => a.EmployeeId == employeeToDelete.Id)
                     .ExecuteDeleteAsync();

                await _context.EmployeeWorkHistories
                     .Where(a => a.EmployeeCode == employeeToDelete.EmployeeCode)
                     .ExecuteDeleteAsync();

                await _context.EmployeeViolation5S
                     .Where(a => a.EmployeeCode == employeeToDelete.EmployeeCode)
                     .ExecuteDeleteAsync();

                await _context.Productivities
                     .Where(a => a.EmployeeCode == employeeToDelete.EmployeeCode)
                     .ExecuteDeleteAsync();

                await _context.Salaries
                     .Where(a => a.EmployeeCode == employeeToDelete.EmployeeCode)
                     .ExecuteDeleteAsync();

                await _context.MonthlyPayroll
                     .Where(a => a.EmployeeCode == employeeToDelete.EmployeeCode)
                     .ExecuteDeleteAsync();

                if (employeeToDelete.Avatar != null)
                {
                    string oldFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "imgs", "avatars", employeeToDelete.Avatar);
                    if (System.IO.File.Exists(oldFilePath))
                    {
                        System.IO.File.Delete(oldFilePath);
                    }
                }

                _context.Employees.Remove(employeeToDelete);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Đã xóa nhân viên và toàn bộ thông tin liên quan!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Index));
        }

        [Authorize(Policy = "ViewEmployee")]
        public async Task<IActionResult> Detail(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var employee = await _context.Employees.AsNoTracking().Where(e => e.Id == id).FirstOrDefaultAsync();

            if (employee == null)
            {
                return NotFound();
            }

            // xử lý phản hồi từ tổ trưởng
            var feedbacks = await _context.Feedbacks.AsNoTracking().Where(f => f.EmployeeId == id && f.Status == 1).OrderByDescending(f => f.Id).ToListAsync();

            ViewBag.Feedbacks = feedbacks;

            // xử lý dữ liệu đào tạo
            var evaluationPeriodsWithScore = await _context.EmployeeTrainingResults
                                                 .AsNoTracking()
                                                 .Where(tr => tr.EmployeeId == id)
                                                 .GroupBy(tr => tr.EvaluationPeriod)
                                                 .Select(g => new
                                                 {
                                                     Id = g.FirstOrDefault().Id,
                                                     EmployeeId = id,
                                                     EvaluationPeriod = g.Key,
                                                     Score = g.FirstOrDefault().Score
                                                 })
                                                 .OrderByDescending(o => o.Id)
                                                 .Take(100)
                                                 .ToListAsync();

            ViewBag.EvaluationPeriods = evaluationPeriodsWithScore;

            var examTrainingAnswers = await _context.ExamTrainingAnswers
                .AsNoTracking()
                .Where(ea => ea.EmployeeCode == employee.EmployeeCode)
                .Join(
                    _context.ExamTrainings,
                    ea => ea.ExamTrainingId,
                    e => e.Id,
                    (ea, e) => new { EmployeeAnswer = ea, Exam = e }
                )
                .Select(joined => new
                {
                    EmplCode = joined.EmployeeAnswer.EmployeeCode,
                    QuizName = joined.Exam.ExamName,
                    ToltalTl = joined.Exam.TlTotal,
                    CountQuestion = _context.QuestionTrainings.Count(q => q.ExamTrainingId == joined.Exam.Id),
                    PointTn = joined.EmployeeAnswer.TnPoint,
                    PointTl = joined.EmployeeAnswer.TlPoint,
                    Id = joined.EmployeeAnswer.Id,
                    Note = joined.EmployeeAnswer.Note
                })
                .OrderByDescending(o => o.Id)
                .Take(100)
                .ToListAsync();

            ViewBag.ExamTrainingAnswers = examTrainingAnswers;

            // xử lý kết quả kiểm tra định kỳ của nhân viên
            var examPeriodicAnswers = await _context.ExamPeriodicAnswers
                .AsNoTracking()
                .Where(ea => ea.EmployeeCode == employee.EmployeeCode)
                .Join(
                    _context.ExamPeriodics,
                    ea => ea.ExamId,
                    e => e.Id,
                    (ea, e) => new { EmployeeAnswer = ea, Exam = e }
                )
                .Select(joined => new
                {
                    EmplCode = joined.EmployeeAnswer.EmployeeCode,
                    QuizName = joined.Exam.ExamName,
                    ToltalTl = joined.Exam.TlTotal,
                    CountQuestion = _context.Questions.Count(q => q.ExamId == joined.Exam.Id),
                    PointTn = joined.EmployeeAnswer.TnPoint,
                    PointTl = joined.EmployeeAnswer.TlPoint,
                    Id = joined.EmployeeAnswer.Id,
                    Note = joined.EmployeeAnswer.Note
                })
                .OrderByDescending(o => o.Id)
                .Take(100)
                .ToListAsync();

            ViewBag.ExamPeriodicAnswers = examPeriodicAnswers;

            //xử lý dữ liệu chạy thử lý thuyết
            var resTrialRun = await (from answer in _context.ExamTrialRunAnswers
                                     join exam in _context.ExamTrialRuns on answer.ExamTrialRunId equals exam.Id
                                     where answer.EmployeeCode == employee.EmployeeCode
                                     select new
                                     {
                                         Id = answer.Id,
                                         Correct = answer.MultipleChoiceCorrect + answer.EssayCorrect,
                                         Incorrect = answer.MultipleChoiceInCorrect + answer.EssayInCorrect,
                                         CriticalFail = answer.MultipleChoiceFail + answer.EssayFail,
                                         Note = answer.Note,
                                         ExamName = exam.ExamName,
                                         TestLevel = exam.TestLevel,
                                     })
                                     .OrderByDescending(o => o.Id)
                                     .Take(100)
                                     .ToListAsync();
            ViewBag.TrialRunAnswer = resTrialRun;

            //xử lý dữ liệu chạy thử thực hành
            var testPractices = await _context.TestPractices.AsNoTracking().Where(t => t.EmployeeId == id).OrderByDescending(o => o.Id).ToListAsync();
            ViewBag.TestPractices = testPractices;

            return View(employee);
        }

        [Authorize(Policy = "ViewEmployee")]
        public async Task<IActionResult> GetTrainingResults(int employeeId, string evaluationPeriod)
        {
            var results = await _context.EmployeeTrainingResults.Where(tr => tr.EmployeeId == employeeId && tr.EvaluationPeriod == evaluationPeriod)
                                                            .Join(_context.Trainings,
                                                                tr => tr.TrainingId,
                                                                t => t.Id,
                                                                (tr, t) => new TrainingResultView
                                                                {
                                                                    TrainingName = t.TrainingName,
                                                                    Type = t.Type,
                                                                    Status = tr.Status
                                                                })
                                                            .ToListAsync();

            var groupedResults = results.GroupBy(r => r.Type).ToDictionary(g => g.Key, g => g.ToList());

            return Json(groupedResults);
        }

        [Authorize(Policy = "ViewEmployee")]
        public async Task<IActionResult> GetTrainings(string type)
        {
            var trainings = await _context.Trainings.Where(t => t.Type == type).ToListAsync();
            return Json(trainings);
        }

        [Authorize(Policy = "DataEmpl")]
        [HttpPost]
        public async Task<IActionResult> SaveTrainingResults([FromBody] TrainingData data)
        {
            var results = data.Results;
            var typeTraining = data.Type;
            if (results == null || !results.Any() || typeTraining == null)
            {
                return Json(new { success = false, message = "Không có dữ liệu để lưu!" });
            }

            var doublePeriod = await _context.EmployeeTrainingResults
                                        .CountAsync(d => d.EmployeeId == results[0].EmployeeId
                                        && d.EvaluationPeriod == results[0].EvaluationPeriod);
            if (doublePeriod > 0)
            {
                return Json(new { success = false, message = "Đợt đào tạo này đã có rồi!" });
            }

            int okCount = results.Count(r => r.Status == 1);
            int total = await _context.Trainings.CountAsync(t => t.Type == typeTraining);
            int score = (int)(okCount * 100.0 / total);
            foreach (var result in results)
            {
                result.Score = score;
                _context.EmployeeTrainingResults.Add(result);
                await _context.SaveChangesAsync();
            }

            return Json(new { success = true, message = "Dữ liệu đào tạo đã được lưu thành công!" });
        }

        [Authorize(Policy = "DataEmpl")]
        public async Task<IActionResult> DeleteDataTraining(long emplId, string period)
        {
            if (emplId == null || period ==null)
            {
                return NotFound();
            }
            try
            {
                await _context.EmployeeTrainingResults
                     .Where(a => a.EmployeeId == emplId
                     && a.EvaluationPeriod == period)
                     .ExecuteDeleteAsync();

                TempData["SuccessMessage"] = "Đã xóa dữ liệu đào tạo!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Detail), new { id = emplId });
        }

        [Authorize(Policy = "DataEmpl")]
        public async Task<IActionResult> DeleteFeedBack(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var fbToDelete = await _context.Feedbacks.FindAsync(id);
            if (fbToDelete == null)
            {
                return NotFound();
            }
            var emplId = fbToDelete.EmployeeId;
            try
            {
                _context.Feedbacks.Remove(fbToDelete);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Đã xóa phản hồi!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Detail), new { id = emplId });
        }

        [Authorize(Policy = "DataEmpl")]
        public async Task<IActionResult> GetTestPracticeDetails(long id)
        {
            var testPractice = await _context.TestPractices
                                           .Include(tp => tp.Details)
                                           .FirstOrDefaultAsync(tp => tp.Id == id);

            if (testPractice == null)
            {
                return NotFound();
            }
            return Json(testPractice);
        }

        [Authorize(Policy = "DataEmpl")]
        [HttpPost]
        public async Task<IActionResult> SavePracticeData([FromBody] TestPractice data)
        {
            if (data == null || !ModelState.IsValid)
            {
                return Json(new { success = false, message = "Dữ liệu không hợp lệ. Vui lòng điền đầy đủ thông tin." });
            }
            bool employeeExists = await _context.Employees.AnyAsync(e => e.Id == data.EmployeeId);
            if (!employeeExists)
            {
                return BadRequest(new { success = false, message = "Mã nhân viên không tồn tại." });
            }
            try
            {
                data.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.TestPractices.Add(data);
                await _context.SaveChangesAsync();

                return Ok(new { success = true, message = "Dữ liệu đã được lưu thành công." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = $"Lỗi khi lưu dữ liệu: {ex.Message}" });
            }
        }

        [Authorize(Policy = "DataEmpl")]
        public async Task<IActionResult> DeleteTestPractice(long id)
        {
            if(id == null)
            {
                return NotFound();
            }
            var tpToDelete = await _context.TestPractices
                                   .Include(tp => tp.Details)
                                   .FirstOrDefaultAsync(tp => tp.Id == id);
            if (tpToDelete == null)
            {
                return NotFound();
            }
            var emplId = tpToDelete.EmployeeId;
            try
            {
                _context.TestPracticeDetails.RemoveRange(tpToDelete.Details);

                _context.TestPractices.Remove(tpToDelete);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Đã xóa dữ liệu chạy thử!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Detail), new { id = emplId });
        }

        [Authorize(Policy = "WorkHistoryEmpl")]
        public async Task<IActionResult> UpdateWorkHistory(long id)
        {
            if(id == null)
            {
                return NotFound();
            }
            var employee = await _context.Employees.AsNoTracking().FirstOrDefaultAsync(e => e.Id == id);
            if (employee == null)
            {
                return NotFound();
            }
            ViewBag.Id = employee.Id;
            ViewBag.EmployeeCode = employee.EmployeeCode;
            ViewBag.EmployeeName = employee.EmployeeName;

            var res = await _context.EmployeeWorkHistories.AsNoTracking().Where(e => e.EmployeeCode == employee.EmployeeCode).ToListAsync();

            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewData["Departments"] = departmentsList;

            return View(res);
        }

        [Authorize(Policy = "WorkHistoryEmpl")]
        [HttpPost]
        public async Task<IActionResult> AddWorkHistory(EmployeeWorkHistory model)
        {
            if (!ModelState.IsValid)
            {
                var errors = ModelState.Values.SelectMany(v => v.Errors).Select(e => e.ErrorMessage);
                return Json(new { success = false, message = "Dữ liệu không hợp lệ!", errors = errors });
            }

            try
            {
                var employee = await _context.Employees
                    .AsNoTracking()
                    .FirstOrDefaultAsync(e => e.EmployeeCode == model.EmployeeCode);

                if (employee == null)
                {
                    return Json(new { success = false, message = "Không tìm thấy thông tin nhân viên." });
                }

                var existingHistories = await _context.EmployeeWorkHistories
                    .AsNoTracking()
                    .Where(h => h.EmployeeCode == model.EmployeeCode)
                    .OrderByDescending(h => h.StartDate)
                    .ToListAsync();

                DateOnly hireDate = employee.HireDate;


                if (model.EndDate.HasValue && model.EndDate.Value <= model.StartDate)
                {
                    return Json(new { success = false, message = "Ngày Kết thúc phải sau Ngày Bắt đầu." });
                }

                if (model.StartDate < hireDate)
                {
                    return Json(new { success = false, message = $"Ngày Bắt đầu phải sau Ngày vào công ty ({hireDate:dd/MM/yyyy})." });
                }

                var openHistory = existingHistories.FirstOrDefault(h => !h.EndDate.HasValue);
                if (openHistory != null)
                {
                    return Json(new { success = false, message = $"Nhân viên đang có quá trình làm việc tại bộ phận '{openHistory.Department}' chưa được kết thúc. Vui lòng cập nhật Ngày Kết thúc trước khi thêm mới." });
                }

                var lastHistory = existingHistories.FirstOrDefault();
                if (lastHistory != null)
                {
                    if (!lastHistory.EndDate.HasValue)
                    {
                    }
                    else if (model.StartDate <= lastHistory.EndDate.Value)
                    {
                        return Json(new { success = false, message = $"Ngày Bắt đầu mới phải sau Ngày Kết thúc của quá trình làm việc gần nhất ({lastHistory.EndDate.Value:dd/MM/yyyy})." });
                    }
                }

                var history = new EmployeeWorkHistory
                {
                    EmployeeCode = model.EmployeeCode,
                    EmployeeName = model.EmployeeName,
                    Department = model.Department,
                    StartDate = model.StartDate,
                    EndDate = model.EndDate,
                    Note = model.Note
                };

                _context.EmployeeWorkHistories.Add(history);
                await _context.SaveChangesAsync();

                return Json(new { success = true, message = "Cập nhật lịch sử công việc thành công!", id = history.Id });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Lỗi hệ thống khi lưu dữ liệu: " + ex.Message });
            }
        }

        [Authorize(Policy = "WorkHistoryEmpl")]
        [HttpPost]
        public async Task<IActionResult> EditWorkHistory(EmployeeWorkHistory model)
        {
            if (!ModelState.IsValid)
            {
                var errors = ModelState.Values.SelectMany(v => v.Errors).Select(e => e.ErrorMessage);
                return Json(new { success = false, message = "Dữ liệu không hợp lệ!", errors = errors });
            }

            var existingHistory = await _context.EmployeeWorkHistories.FindAsync(model.Id);

            if (existingHistory == null)
            {
                return Json(new { success = false, message = "Không tìm thấy dữ liệu." });
            }

            try
            {
                var employee = await _context.Employees
                    .AsNoTracking()
                    .FirstOrDefaultAsync(e => e.EmployeeCode == existingHistory.EmployeeCode);

                if (employee == null)
                {
                    return Json(new { success = false, message = "Không tìm thấy thông tin nhân viên." });
                }

                var allHistories = await _context.EmployeeWorkHistories
                    .AsNoTracking()
                    .Where(h => h.EmployeeCode == existingHistory.EmployeeCode)
                    .OrderBy(h => h.StartDate)
                    .ToListAsync();

                DateOnly hireDate = employee.HireDate;

                if (model.EndDate.HasValue && model.EndDate.Value <= model.StartDate)
                {
                    return Json(new { success = false, message = "Ngày Kết thúc phải sau Ngày Bắt đầu." });
                }

                if (model.StartDate < hireDate)
                {
                    return Json(new { success = false, message = $"Ngày Bắt đầu phải sau Ngày vào công ty ({hireDate:dd/MM/yyyy})." });
                }

                var conflictingHistory = allHistories.FirstOrDefault(h =>
                    h.Id != model.Id &&
                    (
                        (model.StartDate >= h.StartDate && (!h.EndDate.HasValue || model.StartDate < h.EndDate.Value)) ||
                        (h.StartDate >= model.StartDate && (!model.EndDate.HasValue || h.StartDate < model.EndDate.Value))
                    ));

                if (conflictingHistory != null)
                {
                    string conflictingPeriod = conflictingHistory.EndDate.HasValue ?
                        $"{conflictingHistory.StartDate:dd/MM/yyyy} - {conflictingHistory.EndDate.Value:dd/MM/yyyy}" :
                        $"từ {conflictingHistory.StartDate:dd/MM/yyyy} đến hiện tại";

                    return Json(new { success = false, message = $"Khoảng thời gian này bị trùng với quá trình làm việc tại bộ phận '{conflictingHistory.Department}' ({conflictingPeriod})." });
                }

                if (!model.EndDate.HasValue)
                {
                    var nextHistory = allHistories.FirstOrDefault(h =>
                        h.Id != model.Id &&
                        h.StartDate > model.StartDate);

                    if (nextHistory != null)
                    {
                        return Json(new { success = false, message = $"Không thể để quá trình làm việc này mở (không có Ngày Kết thúc) vì đã có quá trình làm việc khác bắt đầu vào {nextHistory.StartDate:dd/MM/yyyy}." });
                    }
                }

                existingHistory.Department = model.Department;
                existingHistory.StartDate = model.StartDate;
                existingHistory.EndDate = model.EndDate;
                existingHistory.Note = model.Note;

                _context.EmployeeWorkHistories.Update(existingHistory);
                await _context.SaveChangesAsync();

                return Json(new { success = true, message = "Cập nhật lịch sử công việc thành công!" });
            }
            catch (Exception ex)
            {
                var errors = ModelState.Values.SelectMany(v => v.Errors).Select(e => e.ErrorMessage);
                return Json(new { success = false, message = "Lỗi khi cập nhật dữ liệu: " + ex.Message, errors = errors });
            }
        }

        [Authorize(Policy = "WorkHistoryEmpl")]
        public async Task<IActionResult> DeleteWorkHistory(long id)
        {
            var history = await _context.EmployeeWorkHistories.FindAsync(id);

            if (history == null)
            {
                return Json(new { success = false, message = "Không tìm thấy dữ liệu." });
            }

            try
            {
                _context.EmployeeWorkHistories.Remove(history);
                await _context.SaveChangesAsync();

                return Json(new { success = true, message = "Xóa dữ liệu thành công!" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Lỗi khi xóa dữ liệu: " + ex.Message });
            }
        }

        [Authorize(Policy = "EmplViolation5S")]
        public async Task<IActionResult> ViewViolation5S(long id)
        {
            if(id == null)
            {
                return NotFound();
            }
            var employee = await _context.Employees.FindAsync(id);
            if (employee == null)
            {
                return NotFound();
            }

            var violation5S = await _context.Violation5S.AsNoTracking().ToListAsync();
            var violation5SList = violation5S.Select(e => new
            {
                Value = e.Id,
                Text = e.Content5S
            });
            ViewBag.ListViolation5S = new SelectList(violation5SList, "Value", "Text");

            var list5S = await _context.EmployeeViolation5S
                                .AsNoTracking()
                                .Where(ev => ev.EmployeeCode == employee.EmployeeCode)
                                .Join(
                                    _context.Violation5S,
                                    ev => ev.Violation5SId,
                                    v => v.Id,
                                    (ev, v) => new EmployeeViolation5SView
                                    {
                                        Id = ev.Id,
                                        EmployeeCode = ev.EmployeeCode,
                                        Content5S = v.Content5S,
                                        DateMonth = ev.DateMonth,
                                        Qty = ev.Qty,
                                        Amount = v.Amount
                                    }
                                )
                                .OrderByDescending(o => o.DateMonth)
                                .ToListAsync();
            ViewBag.EmployeeCode = employee.EmployeeCode;
            ViewBag.EmployeeName = employee.EmployeeName;
            return View(list5S);
        }

        [Authorize(Policy = "EmplViolation5S")]
        [HttpPost]
        public async Task<IActionResult> AddViolation([FromBody] EmployeeViolation5S model)
        {
            var employeeExists = await _context.Employees.AnyAsync(e => e.EmployeeCode == model.EmployeeCode);
            if (!employeeExists)
            {
                return Json(new { success = false, message = "Dữ liệu không hợp lệ!" });
            }
            var violation5S = await _context.Violation5S.AnyAsync(e => e.Id == model.Violation5SId);
            if (!violation5S)
            {
                return Json(new { success = false, message = "Dữ liệu không hợp lệ!" });
            }
            var existingViolation = await _context.EmployeeViolation5S
            .FirstOrDefaultAsync(ev =>
                ev.EmployeeCode == model.EmployeeCode &&
                ev.Violation5SId == model.Violation5SId &&
                ev.DateMonth == model.DateMonth);

            if (existingViolation != null)
            {
                int newQuantity = model.Qty > 0 ? model.Qty : 1;

                existingViolation.Qty += newQuantity;

                try
                {
                    _context.EmployeeViolation5S.Update(existingViolation);
                    await _context.SaveChangesAsync();

                    var successMessage = $"Cập nhật thành công! Đã cộng thêm {newQuantity} lỗi vào bản ghi hiện có. Số lượng lỗi hiện tại: {existingViolation.Qty}";
                    return Json(new { success = true, message = successMessage });
                }
                catch (Exception)
                {
                    return Json(new { success = false, message = "Lỗi khi cập nhật số lần lỗi vào Database!" });
                }
            }
            try
            {
                if (model.Qty <= 0)
                {
                    model.Qty = 1;
                }
                _context.EmployeeViolation5S.Add(model);
                await _context.SaveChangesAsync();
                return Json(new { success = true, message = "Đã lưu dữ liệu!" });

            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Lỗi khi lưu dữ liệu!" });
            }
        }

        [Authorize(Policy = "EmplViolation5S")]
        public async Task<IActionResult> GetViolationDetails(long id)
        {
            var violation = await _context.EmployeeViolation5S.FindAsync(id);
            if (violation == null)
            {
                return NotFound();
            }
            return Json(new
            {
                id = violation.Id,
                violation5SId = violation.Violation5SId,
                dateMonth = violation.DateMonth,
                qty = violation.Qty
            });
        }

        [Authorize(Policy = "EmplViolation5S")]
        [HttpPost]
        public async Task<IActionResult> UpdateViolation([FromBody] EmployeeViolation5S model)
        {
            var empl = await _context.Employees.CountAsync(e => e.EmployeeCode == model.EmployeeCode);
            if (empl == 0)
            {
                return Json(new { success = false, message = "Dữ liệu không hợp lệ!" });
            }
            var violation5S = await _context.Violation5S.CountAsync(v => v.Id == model.Violation5SId);
            if (violation5S == 0)
            {
                return Json(new { success = false, message = "Dữ liệu không hợp lệ!" });
            }
            var isDuplicate = await _context.EmployeeViolation5S
            .AnyAsync(ev =>
                ev.Id != model.Id &&
                ev.EmployeeCode == model.EmployeeCode &&
                ev.Violation5SId == model.Violation5SId &&
                ev.DateMonth == model.DateMonth);

            if (isDuplicate)
            {
                return Json(new { success = false, message = "Nhân viên đã có lỗi trong ngày này vui lòng cập nhật số lần mắc lỗi!" });
            }
            try
            {
                _context.EmployeeViolation5S.Update(model);
                await _context.SaveChangesAsync();
                return Json(new { success = true, message = "Đã lưu dữ liệu!" });

            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Lỗi khi lưu dữ liệu!" });
            }
        }

        [Authorize(Policy = "EmplViolation5S")]
        public async Task<IActionResult> DeleteViolation(long id)
        {
            if(id == null)
            {
                return NotFound();
            }
            var violation5S = await _context.EmployeeViolation5S.FindAsync(id);
            if (violation5S == null)
            {
                return NotFound();
            }
            try
            {
                _context.EmployeeViolation5S.Remove(violation5S);
                await _context.SaveChangesAsync();
                return Json(new { success = true, message = "Xóa lỗi 5S thành công!" });
            }
            catch(Exception ex)
            {
                return Json(new { success = false, message = "Lỗi khi xóa dữ liệu!" });
            }
        }
    }
}
