using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;
using Web_QM.Models.ViewModels;

namespace Web_QM.Controllers
{
    public class DataEmployeeController : Controller
    {
        private readonly QMContext _context;

        private readonly IAuthorizationService _authorizationService;

        public DataEmployeeController( QMContext context, IAuthorizationService authorizationService)
        {
            _context = context;
            _authorizationService = authorizationService;
        }

        [Authorize(Policy = "ClientViewEmpl")]
        public IActionResult Index()
        {
            return View();
        }

        [Authorize(Policy = "ClientViewEmpl")]
        public async Task<IActionResult> GetEmployees(string search, string department)
        {
            var permissions = User.Claims.Where(c => c.Type == "Permission").Select(c => c.Value).ToList();
            var canViewAll = permissions.Any(p => p.Equals("EmployeeAll.View", StringComparison.OrdinalIgnoreCase));
            var canViewDept = permissions.Any(p => p.Equals("EmployeeAllDepartment.View", StringComparison.OrdinalIgnoreCase));
            var userDepartment = User.FindFirst("Department")?.Value;

            var query = _context.Employees.AsNoTracking();

            if (!canViewAll)
            {
                if (string.IsNullOrEmpty(userDepartment))
                {
                    return Json(new { data = new List<object>(), message = "Không có quyền truy cập!" });
                }

                if (canViewDept)
                {
                    var allChildDepts = await GetChildDepartments(userDepartment);
                    query = query.Where(e => allChildDepts.Contains(e.Department));
                }
                else
                {
                    query = query.Where(e => e.Department == userDepartment);
                }
            }

            if (!string.IsNullOrEmpty(search))
            {
                var lowerSearch = search.Trim().ToLower();
                query = query.Where(e => e.EmployeeName.ToLower().Contains(lowerSearch) || e.EmployeeCode.ToLower().Contains(lowerSearch));
            }

            if (canViewAll && !string.IsNullOrEmpty(department))
            {
                query = query.Where(e => e.Department == department);
            }

            var employees = await query.OrderBy(o => o.EmployeeCode).ToListAsync();
            return Json(new { data = employees });
        }

        private async Task<List<string>> GetChildDepartments(string deptName)
        {
            var result = new List<string> { deptName };
            var allDepts = await _context.Departments.AsNoTracking().ToListAsync();
            FindChildren(deptName, allDepts, result);
            return result;
        }

        private void FindChildren(string parentName, List<Department> allDepts, List<string> result)
        {
            var children = allDepts.Where(d => d.DpParent == parentName).Select(d => d.DepartmentName).ToList();
            foreach (var child in children)
            {
                if (!result.Contains(child))
                {
                    result.Add(child);
                    FindChildren(child, allDepts, result);
                }
            }
        }

        [Authorize(Policy = "ClientViewEmpl")]
        public async Task<IActionResult> GetTopEmployeeProductivity(int year)
        {
            if (year < 2020 || year > 2100)
            {
                return BadRequest(new { message = "Năm không hợp lệ." });
            }
            var groupedData = await _context.Productivities
                .AsNoTracking()
                .Where(p => p.MeasurementYear == year)
                .Join(
                    _context.Employees,
                    p => p.EmployeeCode,
                    e => e.EmployeeCode,
                    (p, e) => new { Productivity = p, Employee = e }
                )
                .GroupBy(x => new {
                    x.Employee.EmployeeCode,
                    x.Employee.EmployeeName,
                    x.Employee.Department
                })
                .Select(g => new
                {
                    g.Key.EmployeeCode,
                    g.Key.EmployeeName,
                    g.Key.Department,
                    TotalScore = g.Sum(x => x.Productivity.ProductivityScore),
                    Count = g.Count()
                })
                .OrderByDescending(x => x.TotalScore)
                .Take(10)
                .ToListAsync();

            if (!groupedData.Any())
            {
                return NotFound(new { message = $"Không tìm thấy dữ liệu năng suất cho năm {year}." });
            }
            var resultList = groupedData
                .Select((data, index) => new
                {
                    Rank = index + 1,
                    EmployeeCode = data.EmployeeCode,
                    EmployeeName = data.EmployeeName,
                    Department = data.Department,
                    AverageScore = data.Count > 0 ? data.TotalScore / data.Count : 0,
                })
                .ToList();

            return Ok(resultList);
        }

        [Authorize(Policy = "ClientViewEmpl")]
        public async Task<IActionResult> GetTopEmployeeKaizen(int year)
        {
            if (year < 2020 || year > 2100)
            {
                return BadRequest(new { message = "Năm không hợp lệ." });
            }
            var allKaizenData = await _context.Kaizens
                .AsNoTracking()
                .Where(k => k.DateMonth.Year == year)
                .Join(
                    _context.Employees.AsNoTracking(),
                    k => k.EmployeeCode,
                    e => e.EmployeeCode,
                    (k, e) => new
                    {
                        k.EmployeeCode,
                        e.EmployeeName,
                        e.Department,
                        KaizenDate = k.DateMonth,
                        k.ManagementReview
                    }
                )
                .ToListAsync();

            if (!allKaizenData.Any())
            {
                return NotFound(new { message = $"Không tìm thấy dữ liệu Kaizen cho năm {year}." });
            }

            int GetReviewPriority(string review)
            {
                return review?.ToUpper() switch
                {
                    "A" => 4,
                    "B" => 3,
                    "C" => 2,
                    "D" => 1,
                    _ => 0,
                };
            }
            var groupedData = allKaizenData
                .GroupBy(x => new { x.EmployeeCode, x.EmployeeName, x.Department })
                .Select(g => new
                {
                    g.Key.EmployeeCode,
                    g.Key.EmployeeName,
                    g.Key.Department,
                    KaizenCount = g.Count(),
                    TotalPriorityScore = g.Sum(x => GetReviewPriority(x.ManagementReview)),
                    EarliestKaizenDate = g.Min(x => x.KaizenDate)
                })
                .OrderByDescending(x => x.KaizenCount)
                .ThenByDescending(x => x.TotalPriorityScore)
                .ThenBy(x => x.EarliestKaizenDate)
                .Take(10)
                .ToList();
            var resultList = groupedData
                .Select((data, index) => new
                {
                    Rank = index + 1,
                    EmployeeCode = data.EmployeeCode,
                    EmployeeName = data.EmployeeName,
                    Department = data.Department,
                    TotalScore = data.KaizenCount,
                })
                .ToList();

            return Ok(resultList);
        }

        [Authorize(Policy = "ClientViewEmpl")]
        public async Task<IActionResult> GetTopEmployeeErrors(int year)
        {
            if (year < 2020 || year > 2100)
            {
                return BadRequest(new { message = "Năm không hợp lệ." });
            }

            var errorDataCounts = await _context.ErrorDatas
                .AsNoTracking()
                .Where(e => e.DateMonth.Year == year)
                .GroupBy(e => e.EmployeeCode)
                .Select(g => new
                {
                    EmployeeCode = g.Key,
                    ProductionErrorCount = g.Count()
                })
                .ToListAsync();

            var violation5SCounts = await _context.EmployeeViolation5S
                .AsNoTracking()
                .Where(v => v.DateMonth.Year == year)
                .GroupBy(v => v.EmployeeCode)
                .Select(g => new
                {
                    EmployeeCode = g.Key,
                    Violation5SCount = g.Count()
                })
                .ToListAsync();

            var allEmployeeCodes = errorDataCounts
                .Select(x => x.EmployeeCode)
                .Union(violation5SCounts.Select(x => x.EmployeeCode))
                .ToList();

            var combinedErrorData = allEmployeeCodes.Select(code =>
            {
                var prodError = errorDataCounts.FirstOrDefault(x => x.EmployeeCode == code);
                var violation5S = violation5SCounts.FirstOrDefault(x => x.EmployeeCode == code);

                int prodCount = prodError?.ProductionErrorCount ?? 0;
                int v5sCount = violation5S?.Violation5SCount ?? 0;
                int totalCount = prodCount + v5sCount;

                return new
                {
                    EmployeeCode = code,
                    ProductionErrorCount = prodCount,
                    Violation5SCount = v5sCount,
                    TotalErrorCount = totalCount
                };
            })
            .OrderByDescending(x => x.TotalErrorCount)
            .ToList();

            if (!combinedErrorData.Any())
            {
                return NotFound(new { message = $"Không tìm thấy dữ liệu lỗi cho năm {year}." });
            }

            var allEmployees = await _context.Employees.AsNoTracking().ToListAsync();

            var topEmployeeErrorDetails = allEmployees
            .Join(
                combinedErrorData,
                e => e.EmployeeCode,
                c => c.EmployeeCode,
                (e, c) => new
                {
                    EmployeeCode = e.EmployeeCode,
                    EmployeeName = e.EmployeeName,
                    Department = e.Department,
                    ProductionErrorCount = c.ProductionErrorCount,
                    Violation5SCount = c.Violation5SCount,
                    TotalErrorCount = c.TotalErrorCount
                }
            )
            .OrderByDescending(x => x.TotalErrorCount)
            .Take(10)
            .ToList();
            var resultList = topEmployeeErrorDetails
                .Select((data, index) => new
                {
                    Rank = index + 1,
                    data.EmployeeCode,
                    data.EmployeeName,
                    data.Department,
                    data.ProductionErrorCount,
                    data.Violation5SCount,
                    data.TotalErrorCount
                })
                .ToList();

            return Ok(resultList);
        }

        [Authorize(Policy = "ClientViewEmpl")]
        public async Task<IActionResult> Detail(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var employee = await _context.Employees
                    .AsNoTracking()
                    .Where(e => e.Id == id)
                    .FirstOrDefaultAsync();

            if (employee == null)
            {
                return NotFound();
            }

            var permissions = User.Claims.Where(c => c.Type == "Permission").Select(c => c.Value).ToList();
            var canViewAll = permissions.Any(p => p.Equals("EmployeeAll.View", StringComparison.OrdinalIgnoreCase));
            if (!canViewAll)
            {
                var userDepartment = User.FindFirst("Department")?.Value;
                if (userDepartment == null || employee.Department != userDepartment)
                {
                    return NotFound();
                }
            }

            var authAddFeedback = await _authorizationService.AuthorizeAsync(User, "ClientAddFeedbackEmpl");
            ViewBag.CanManageFeedback = authAddFeedback.Succeeded;

            // xử lý dữ liệu đào tạo
            var evaluationPeriodsWithScore = await _context.EmployeeTrainingResults
                                                 .AsNoTracking()
                                                 .Where(tr => tr.EmployeeId == id)
                                                 .GroupBy(tr => tr.EvaluationPeriod)
                                                 .Select(g => new
                                                 {
                                                     Id = g.FirstOrDefault().Id,
                                                     EvaluationPeriod = g.Key,
                                                     Score = g.FirstOrDefault().Score
                                                 })
                                                 .OrderByDescending(o => o.Id)
                                                 .Take(100)
                                                 .ToListAsync();

            ViewBag.EvaluationPeriods = evaluationPeriodsWithScore;

            var examTrainingAnswers = await _context.ExamTrainingAnswers
               .AsNoTracking()
               .Where(ea => ea.EmployeeCode == employee.EmployeeCode && ea.IsShow == 1)
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

            // xử lý kết quả kiểm tra của nhân viên
            var examPeriodicAnswers = await _context.ExamPeriodicAnswers
                .AsNoTracking()
                .Where(ea => ea.EmployeeCode == employee.EmployeeCode && ea.IsShow == 1)
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
                    Note = joined.EmployeeAnswer.Note,
                })
                .OrderByDescending(o => o.Id)
                .Take(100)
                .ToListAsync();
            ViewBag.ExamPeriodicAnswers = examPeriodicAnswers;

            //xử lý dữ liệu chạy thử lý thuyết
            var resTrialRun = await (from answer in _context.ExamTrialRunAnswers
                                     join exam in _context.ExamTrialRuns on answer.ExamTrialRunId equals exam.Id
                                     where answer.EmployeeCode == employee.EmployeeCode && answer.IsShow == 1
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
                                     .AsNoTracking()
                                     .OrderByDescending(o => o.Id)
                                     .Take(100)
                                     .ToListAsync();
            ViewBag.TrialRunAnswer = resTrialRun;

            //xử lý dữ liệu chạy thử thực hành
            var testPractices = await _context.TestPractices.AsNoTracking().Where(t => t.EmployeeId == id).OrderByDescending(o => o.Id).Take(100).ToListAsync();
            ViewBag.TestPractices = testPractices;

            //xử lý quá trình làm việc
            var history = await _context.EmployeeWorkHistories.AsNoTracking().Where(e => e.EmployeeCode == employee.EmployeeCode).OrderByDescending(o => o.StartDate).Take(50).ToListAsync();
            ViewBag.WorkHistory = history;

            var processedHistory = history.Select(h => new WorkHistoryView
            {
                StartDate = h.StartDate,
                EndDate = h.EndDate,
                Department = h.Department,
                KaizenCount = 0,
                ErrorCount = 0,
                Violation5SCount = 0
            }).ToList();

            var allKaizens = await _context.Kaizens
                .Where(k => k.EmployeeCode == employee.EmployeeCode)
                .Select(k => k.DateMonth)
                .ToListAsync();

            var allErrorRecords = await _context.ErrorDatas
                .Where(e => e.EmployeeCode == employee.EmployeeCode)
                .Select(e => e.DateMonth)
                .ToListAsync();

            var allViolation5S = await _context.EmployeeViolation5S
                .Where(v => v.EmployeeCode == employee.EmployeeCode)
                .Select(v => new {
                    v.DateMonth,
                    v.Qty
                })
                .ToListAsync();

            foreach (var item in processedHistory)
            {
                DateOnly actualEndDate = item.EndDate.GetValueOrDefault(DateOnly.FromDateTime(DateTime.Today));
                item.KaizenCount = allKaizens
                    .Count(kDate => kDate >= item.StartDate && kDate <= actualEndDate);
                item.ErrorCount = allErrorRecords
                    .Count(eDate => eDate >= item.StartDate && eDate <= actualEndDate);
                item.Violation5SCount = allViolation5S
                    .Where(v => v.DateMonth >= item.StartDate && v.DateMonth <= actualEndDate)
                    .Sum(v => v.Qty);
            }
            ViewBag.WorkHistory = processedHistory;

            return View(employee);
        }

        [Authorize(Policy = "ClientAddFeedbackEmpl")]
        [HttpPost]
        public async Task<IActionResult> SaveFeedback([FromBody]Feedback feedback)
        {
            if (feedback.Comment == null)
            {
                return Json(new { success = false, message = "Nội dung phản hồi không được để trống." });
            }
            try
            {
                var employeeCodeIsLogin = User.FindFirst("EmployeeCode")?.Value;
                var employeeNameIsLogin = User.FindFirst("EmployeeName")?.Value;

                feedback.FeedbackerName = employeeCodeIsLogin + "-" + employeeNameIsLogin;
                feedback.CreatedDate = DateTime.Now.ToString("dd/MM/yyyy");
                feedback.Status = 0;

                var empl = await _context.Employees.FirstOrDefaultAsync(e => e.Id == feedback.EmployeeId);
                if(empl == null)
                {
                    return Json(new { success = false, message = "Nhân viên không tồn tại!" });
                }
                _context.Feedbacks.Add(feedback);
                await _context.SaveChangesAsync();

                return Json(new { success = true, message = "Thêm dữ liệu thành công!" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }

        [Authorize(Policy = "ClientViewEmpl")]
        public async Task<IActionResult> GetFeedbacks(long employeeId)
        {
            if (employeeId == 0)
            {
                return BadRequest();
            }
            var empl = await _context.Employees.FirstOrDefaultAsync(e => e.Id == employeeId);
            if(empl == null)
            {
                return BadRequest();
            }
            IQueryable<Feedback> query = _context.Feedbacks.Where(f => f.EmployeeId == employeeId);
            var authAddFeedback = await _authorizationService.AuthorizeAsync(User, "ClientAddFeedbackEmpl");
            var canAddFeedback = authAddFeedback.Succeeded;

            var emplCodeIsLogin = User.FindFirst("EmployeeCode")?.Value;
            var emplNameIsLogin = User.FindFirst("EmployeeName")?.Value;
            var feedbacker = emplCodeIsLogin + "-" + emplNameIsLogin;   
            if (canAddFeedback)
            {
                query = query.Where(f => f.FeedbackerName == feedbacker || f.Status == 1);
            }
            else
            {
                query = query.Where(f => f.Status == 1);
            }
            var feedbacks = await query.OrderByDescending(f => f.Id).ToListAsync();
            return Json(feedbacks);
        }

        [Authorize(Policy = "ClientAddFeedbackEmpl")]
        [HttpPost]
        public async Task<IActionResult> Approve(long id)
        {
            var feedback = await _context.Feedbacks.FirstOrDefaultAsync(f => f.Id == id);
            if (feedback == null) return NotFound();

            var empl = await _context.Employees.FirstOrDefaultAsync(e => e.Id == feedback.EmployeeId);
            if (empl == null)
            {
                return Json(new { success = false, message = "Nhân viên không tồn tại!" });
            }
            try
            {
                feedback.Status = 1;
                _context.Feedbacks.Update(feedback);
                Notification nt = new Notification()
                {
                    Message = "Nhân viên " + empl.EmployeeCode + "-" + empl.EmployeeName + " có phản hồi mới",
                    IsRead = 0,
                    CreatedDate = DateTime.Now.ToString("dd/MM/yyyy")
                };
                _context.Notifications.Add(nt);

                await _context.SaveChangesAsync();
                return Json(new { success = true, message = "Đã gửi phản hồi!" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }

        [Authorize(Policy = "ClientAddFeedbackEmpl")]
        [HttpGet]
        public async Task<IActionResult> GetFeedbackDetails(long id)
        {
            var feedback = await _context.Feedbacks.FirstOrDefaultAsync(f => f.Id == id);
            if (feedback == null) return NotFound();

            return Json(feedback);
        }

        [Authorize(Policy = "ClientAddFeedbackEmpl")]
        [HttpPost]
        public async Task<IActionResult> EditFeedback([FromBody]Feedback model)
        {
            if (model == null) return BadRequest();

            if (model.Comment == null)
            {
                return Json(new { success = false, message = "Nội dung phản hồi không được để trống." });
            }

            var feedback = await _context.Feedbacks.FirstOrDefaultAsync(f => f.Id == model.Id);
            if (feedback == null) return NotFound();

            try
            {
                var employeeCodeIsLogin = User.FindFirst("EmployeeCode")?.Value;
                var employeeNameIsLogin = User.FindFirst("EmployeeName")?.Value;
                feedback.FeedbackerName = employeeCodeIsLogin + "-" + employeeNameIsLogin;

                feedback.Comment = model.Comment;

                _context.Feedbacks.Update(feedback);
                await _context.SaveChangesAsync();
                return Json(new { success = true, message = "Cập nhật thành công." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }

        [Authorize(Policy = "ClientAddFeedbackEmpl")]
        [HttpPost]
        public async Task<IActionResult> DeleteFeedback(long id)
        {
            var feedback = await _context.Feedbacks.FirstOrDefaultAsync(f => f.Id == id);
            if (feedback == null) return NotFound();

            try
            {
                _context.Feedbacks.Remove(feedback);
                await _context.SaveChangesAsync();
                return Json(new { success = true, message = "Xóa thành công." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Lỗi server: {ex.Message}" });
            }
        }

        [Authorize(Policy = "ClientViewEmpl")]
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

        [Authorize(Policy = "ClientViewEmpl")]
        public async Task<IActionResult> GetTrainings(string type)
        {
            var trainings = await _context.Trainings.Where(t => t.Type == type).ToListAsync();
            return Json(trainings);
        }

        [Authorize(Policy = "ClientViewEmpl")]
        public IActionResult GetProductivityData(string employeeCode, int selectedYear)
        {
            var monthlyProductivity = _context.Productivities
                .Where(p => p.EmployeeCode == employeeCode && p.MeasurementYear == selectedYear)
                .ToList();

            var result = Enumerable.Range(1, 12).Select(month =>
            {
                var data = monthlyProductivity.FirstOrDefault(p => p.MeasurementMonth == month);

                string monthName = $"T{month}";
                decimal score = data?.ProductivityScore ?? 0;

                return new
                {
                    monthLabel = monthName,
                    score = score
                };
            }).ToList();

            return Json(result);
        }

        [Authorize(Policy = "ClientViewEmpl")]
        public IActionResult GetChartSawing(string employeeCode, int selectedYear)
        {
            var monthlyProductivity = _context.SawingPerformances
                .Where(p => p.EmployeeCode == employeeCode && p.MeasurementYear == selectedYear)
                .ToList();

            var result = Enumerable.Range(1, 12).Select(month =>
            {
                var data = monthlyProductivity.FirstOrDefault(p => p.MeasurementMonth == month);

                string monthName = $"T{month}";
                decimal score = data?.SalesRate ?? 0;

                return new
                {
                    monthLabel = monthName,
                    score = score
                };
            }).ToList();

            return Json(result);
        }

        [Authorize(Policy = "ClientViewEmpl")]
        public async Task<IActionResult> GetActivityDetails(string employeeCode, string type, string startDate, string endDate,
                        int? selectedYear)
        {
            DateOnly start = DateOnly.Parse(startDate);
            DateOnly end = DateOnly.Parse(endDate);

            object details = null;

            switch (type)
            {
                case "Kaizen":
                    var queryKaizen = _context.Kaizens
                        .AsNoTracking()
                        .Where(ev => ev.EmployeeCode == employeeCode && ev.DateMonth >= start && ev.DateMonth <= end);

                    if (selectedYear.HasValue)
                    {
                        queryKaizen = queryKaizen.Where(ev => ev.DateMonth.Year == selectedYear.Value);
                    }
                    details = await queryKaizen.OrderByDescending(k => k.DateMonth).ToListAsync();
                    break;

                case "Error":
                    var queryError = _context.ErrorDatas
                        .AsNoTracking()
                        .Where(ev => ev.EmployeeCode == employeeCode && ev.DateMonth >= start && ev.DateMonth <= end);

                    if (selectedYear.HasValue)
                    {
                        queryError = queryError.Where(ev => ev.DateMonth.Year == selectedYear.Value);
                    }

                    details = await queryError.OrderByDescending(e => e.DateMonth).ToListAsync();
                    break;

                case "Violation5S":
                    var query = _context.EmployeeViolation5S
                        .AsNoTracking()
                        .Where(ev => ev.EmployeeCode == employeeCode && ev.DateMonth >= start && ev.DateMonth <= end);

                    if (selectedYear.HasValue)
                    {
                        query = query.Where(ev => ev.DateMonth.Year == selectedYear.Value);
                    }

                    details = await query
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
                            }
                        )
                        .OrderByDescending(o => o.DateMonth)
                        .ToListAsync();
                    break;
            }
            return Json(details);
        }

        [Authorize(Policy = "ClientViewEmpl")]
        public async Task<IActionResult> GetExamPeriodicDetail(long id)
        {
            var answerRecord = await _context.ExamPeriodicAnswers.FindAsync(id);
            if (answerRecord == null) return NotFound();

            var questions = await _context.Questions
                .Where(q => q.ExamId == answerRecord.ExamId)
                .OrderBy(q => q.DisplayOrder)
                .ToListAsync();

            var dictAnswers = (answerRecord.ListAnswer ?? "")
                .Split('-', StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Split('.'))
                .Where(x => x.Length == 2)
                .ToDictionary(x => x[0].Trim(), x => x[1].Trim());

            var questionList = questions.Select((q, index) => {
                var stt = (index + 1).ToString();
                return new
                {
                    stt = index + 1,
                    questionText = q.QuestionText,
                    optionA = q.OptionA,
                    optionB = q.OptionB,
                    optionC = q.OptionC,
                    optionD = q.OptionD,
                    correctOption = q.CorrectOption,
                    employeeChoice = dictAnswers.ContainsKey(stt) ? dictAnswers[stt] : ""
                };
            }).ToList();

            return Json(new
            {
                details = questionList,
                pdfPath = answerRecord.EssayResultPDF
            });
        }

        [Authorize(Policy = "ClientViewEmpl")]
        public async Task<IActionResult> ExportToExcel(long id)
        {
            if (id == null)
            {
                return NotFound();
            }
            var employee = await _context.Employees.FirstOrDefaultAsync(e => e.Id == id);
            if (employee == null)
            {
                return NotFound();
            }
            string employeeName = employee.EmployeeCode + "-" + employee.EmployeeName;

            var productivities = await _context.Productivities
                .AsNoTracking()
                .Where(p => p.EmployeeCode == employee.EmployeeCode)
                .OrderByDescending(p => p.MeasurementYear)
                .ThenByDescending(p => p.MeasurementMonth)
                .ToListAsync();

            var sawingPerformances = await _context.SawingPerformances
                .AsNoTracking()
                .Where(s => s.EmployeeCode == employee.EmployeeCode)
                .OrderByDescending(p => p.MeasurementYear)
                .ThenByDescending(p => p.MeasurementMonth)
                .ToListAsync();

            var kaizens = await _context.Kaizens
                .AsNoTracking()
                .Where(k => k.EmployeeCode == employee.EmployeeCode)
                .OrderByDescending(k => k.DateMonth)
                .ToListAsync();

            var errorDatas = await _context.ErrorDatas
                .AsNoTracking()
                .Where(e => e.EmployeeCode == employee.EmployeeCode)
                .OrderByDescending(e => e.DateMonth)
                .ToListAsync();

            var violation5S = await _context.EmployeeViolation5S
                .AsNoTracking()
                .Where(v => v.EmployeeCode == employee.EmployeeCode)
                .Join(
                    _context.Violation5S,
                    ev => ev.Violation5SId,
                    v => v.Id,
                    (ev, v) => new
                    {
                        Content5S = v.Content5S,
                        DateMonth = ev.DateMonth,
                        Qty = ev.Qty,
                    }
                )
                .OrderByDescending(v => v.DateMonth)
                .ToListAsync();

            var trainingresults = await _context.EmployeeTrainingResults
                .AsNoTracking()
                .Where(etr => etr.EmployeeId == id)
                .GroupBy(etr => etr.EvaluationPeriod)
                .Select(g => g.OrderBy(x => x.Id).FirstOrDefault())
                .ToListAsync();

            var examTrainingAnswers = await _context.ExamTrainingAnswers
                .AsNoTracking()
                .Where(a => a.EmployeeCode == employee.EmployeeCode)
                .Join(_context.ExamTrainings,
                    answer => answer.ExamTrainingId,
                    exam => exam.Id,
                    (answer, exam) => new { answer, exam })
                .Select(x => new
                {
                    ExamName = x.exam.ExamName,
                    Total = x.exam.TlTotal + _context.QuestionTrainings.Count(q => q.ExamTrainingId == x.answer.ExamTrainingId),
                    Answer = x.answer.TnPoint + x.answer.TlPoint,
                })
                .ToListAsync();

            var examPeriodicAnswers = await _context.ExamPeriodicAnswers
                .AsNoTracking()
                .Where(a => a.EmployeeCode == employee.EmployeeCode)
                .Join(_context.ExamPeriodics,
                    answer => answer.ExamId,
                    exam => exam.Id,
                    (answer, exam) => new { answer, exam })
                .Select(x => new
                {
                    ExamName = x.exam.ExamName,
                    Total = x.exam.TlTotal + _context.Questions.Count(q => q.ExamId == x.answer.ExamId),
                    Answer = x.answer.TnPoint + x.answer.TlPoint,
                })
                .ToListAsync();

            var examTrialRunAnswers = await _context.ExamTrialRunAnswers
                .AsNoTracking()
                .Where(a => a.EmployeeCode == employee.EmployeeCode)
                .Join(_context.ExamTrialRuns,
                    answer => answer.ExamTrialRunId,
                    exam => exam.Id,
                    (answer, exam) => new { answer, exam })
                .Select(x => new
                {
                    ExamName = x.exam.ExamName,
                    Correct = x.answer.MultipleChoiceCorrect + x.answer.EssayCorrect,
                    Incorrect = x.answer.MultipleChoiceInCorrect + x.answer.EssayInCorrect,
                    CriticalFail = x.answer.MultipleChoiceFail + x.answer.EssayFail,
                })
                .ToListAsync();

            var testPractices = await _context.TestPractices
                .AsNoTracking()
                .Where(tp => tp.EmployeeId == id)
                .Select(tp => new
                {
                    tp.TestName,
                    tp.TestLevel,
                    tp.PartName,
                    tp.Result
                })
                .ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                int row;
                int stt = 1;
                int lastRow;

                var ws1 = workbook.Worksheets.Add("Chung");
                ws1.Cell(1, 1).Value = "Mã nhân viên";
                ws1.Cell(1, 2).Value = "Tên nhân viên";
                ws1.Cell(1, 3).Value = "Ngày sinh";
                ws1.Cell(1, 4).Value = "Giới tính";
                ws1.Cell(1, 5).Value = "Bộ phận";
                ws1.Cell(1, 6).Value = "Ngày vào công ty";
                ws1.Cell(1, 7).Value = "Vị trí";
                ws1.Cell(1, 8).Value = "Ghi chú";

                ws1.Range("A1:H1").Style.Font.Bold = true;
                ws1.Range("A1:H1").Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF1F5");
                ws1.Range("A1:H1").Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
                ws1.Range("A1:H1").Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);

                ws1.Cell(2, 1).Value = employee.EmployeeCode;
                ws1.Cell(2, 2).Value = employee.EmployeeName;
                ws1.Cell(2, 3).Value = employee.DateOfBirth.ToString();
                ws1.Cell(2, 4).Value = employee.Gender;
                ws1.Cell(2, 5).Value = employee.Department;
                ws1.Cell(2, 6).Value = employee.HireDate.ToString();
                ws1.Cell(2, 7).Value = employee.Position;
                ws1.Cell(2, 8).Value = employee.Note;
                ws1.Range("A2:H2").Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws1.Columns().AdjustToContents();

                if (productivities.Count > 0)
                {
                    var ws_Productivity = workbook.Worksheets.Add("Năng suất");
                    ws_Productivity.Cell(1, 1).Value = "Tháng/Năm";
                    ws_Productivity.Cell(1, 2).Value = "Năng suất";
                    ws_Productivity.Range("A1:B1").Style.Font.Bold = true;
                    ws_Productivity.Range("A1:B1").Style.Fill.BackgroundColor = XLColor.LightYellow;
                    ws_Productivity.SheetView.FreezeRows(1);
                    foreach (var productivity in productivities)
                    {
                        row = ws_Productivity.LastRowUsed().RowNumber() + 1;
                        ws_Productivity.Cell(row, 1).Value = $"{productivity.MeasurementMonth}/{productivity.MeasurementYear}";
                        ws_Productivity.Cell(row, 2).Value = $"{productivity.ProductivityScore} %";

                        decimal score;
                        if (decimal.TryParse(productivity.ProductivityScore.ToString(), out score))
                        {
                            if (score >= 90)
                            {
                                ws_Productivity.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.Green;
                            }
                            else if (score < 70)
                            {
                                ws_Productivity.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.Red;
                            }
                        }
                    }
                    lastRow = ws_Productivity.LastRowUsed().RowNumber();
                    if (lastRow > 1)
                    {
                        var table = ws_Productivity.Range(1, 1, lastRow, 2).CreateTable();
                        table.Theme = XLTableTheme.None;
                    }
                    ws_Productivity.Columns().AdjustToContents();
                }

                if (sawingPerformances.Count > 0)
                {
                    var ws_Sawing = workbook.Worksheets.Add("Doanh số cưa");
                    ws_Sawing.Cell(1, 1).Value = "Tháng/Năm";
                    ws_Sawing.Cell(1, 2).Value = "Doanh số cưa (USD)";
                    ws_Sawing.Cell(1, 3).Value = "Thời gian làm việc (phút)";
                    ws_Sawing.Cell(1, 4).Value = "Doanh số/Thời gian";
                    ws_Sawing.Range("A1:D1").Style.Font.Bold = true;
                    ws_Sawing.Range("A1:D1").Style.Fill.BackgroundColor = XLColor.LightYellow;
                    ws_Sawing.SheetView.FreezeRows(1);
                    foreach (var sawingPerformance in sawingPerformances)
                    {
                        row = ws_Sawing.LastRowUsed().RowNumber() + 1;
                        ws_Sawing.Cell(row, 1).Value = $"{sawingPerformance.MeasurementMonth}/{sawingPerformance.MeasurementYear}";
                        ws_Sawing.Cell(row, 2).Value = sawingPerformance.SalesAmountUSD;
                        ws_Sawing.Cell(row, 3).Value = sawingPerformance.WorkMinute;
                        ws_Sawing.Cell(row, 4).Value = sawingPerformance.SalesRate;
                    }
                    lastRow = ws_Sawing.LastRowUsed().RowNumber();
                    if (lastRow > 1)
                    {
                        var table = ws_Sawing.Range(1, 1, lastRow, 4).CreateTable();
                        table.Theme = XLTableTheme.None;
                    }
                    ws_Sawing.Columns().AdjustToContents();
                }

                var ws3 = workbook.Worksheets.Add("Kaizen");
                ws3.Cell(1, 1).Value = "STT";
                ws3.Cell(1, 2).Value = "Tháng/Năm";
                ws3.Cell(1, 3).Value = "Mục tiêu cải tiến";
                ws3.Cell(1, 4).Value = "Tiêu đề cải tiến";
                ws3.Cell(1, 5).Value = "Tình trạng hiện tại";
                ws3.Cell(1, 6).Value = "Ý kiến cài tiến";
                ws3.Cell(1, 7).Value = "Hiệu quả, lợi ích";
                ws3.Cell(1, 8).Value = "Tổ trưởng đánh giá";
                ws3.Cell(1, 9).Value = "BQL đánh giá";
                ws3.Cell(1, 10).Value = "Thời hạn hoàn thành";
                ws3.Cell(1, 11).Value = "Thời gian bắt đầu sử dụng";
                ws3.Cell(1, 12).Value = "Tình trạng hiện tại (sau kaizen)";
                ws3.Cell(1, 13).Value = "Ghi chú";
                ws3.Range("A1:N1").Style.Font.Bold = true;
                ws3.Range("A1:N1").Style.Fill.BackgroundColor = XLColor.FromHtml("#D9E1F2");
                ws3.SheetView.FreezeRows(1);
                stt = 1;
                foreach (var kaizen in kaizens)
                {
                    row = ws3.LastRowUsed().RowNumber() + 1;
                    ws3.Cell(row, 1).Value = stt++;
                    ws3.Cell(row, 2).Value = kaizen.DateMonth.ToString();
                    ws3.Cell(row, 3).Value = kaizen.ImprovementGoal;
                    ws3.Cell(row, 4).Value = kaizen.ImprovementTitle;
                    ws3.Cell(row, 5).Value = kaizen.CurrentSituation;
                    ws3.Cell(row, 5).Style.Alignment.WrapText = true;
                    ws3.Cell(row, 6).Value = kaizen.ProposedIdea;
                    ws3.Cell(row, 6).Style.Alignment.WrapText = true;
                    ws3.Cell(row, 7).Value = kaizen.EstimatedBenefit;
                    ws3.Cell(row, 7).Style.Alignment.WrapText = true;
                    ws3.Cell(row, 8).Value = kaizen.TeamLeaderRating;
                    ws3.Cell(row, 9).Value = kaizen.ManagementReview;
                    ws3.Cell(row, 10).Value = kaizen.Deadline;
                    ws3.Cell(row, 11).Value = kaizen.StartTime;
                    ws3.Cell(row, 12).Value = kaizen.CurrentStatus;
                    ws3.Cell(row, 13).Value = kaizen.Note;
                }
                lastRow = ws3.LastRowUsed().RowNumber();
                if (lastRow > 1)
                {
                    var table = ws3.Range(1, 1, lastRow, 14).CreateTable();
                    table.Theme = XLTableTheme.None;
                }
                ws3.Columns().AdjustToContents();

                var ws4 = workbook.Worksheets.Add("Lỗi sản xuất");
                ws4.Cell(1, 1).Value = "STT";
                ws4.Cell(1, 2).Value = "Ngày/Tháng";
                ws4.Cell(1, 3).Value = "Order";
                ws4.Cell(1, 4).Value = "Tên chi tiết";
                ws4.Cell(1, 5).Value = "Số lượng order";
                ws4.Cell(1, 6).Value = "Số lượng NG";
                ws4.Cell(1, 7).Value = "Tỷ lệ NG";
                ws4.Cell(1, 8).Value = "Phát hiện lỗi";
                ws4.Cell(1, 9).Value = "Dạng lỗi";
                ws4.Cell(1, 10).Value = "Nguyên nhân lỗi";
                ws4.Cell(1, 11).Value = "Nội dung lỗi";
                ws4.Cell(1, 12).Value = "Nhận định dung sai";
                ws4.Cell(1, 13).Value = "Nguyên nhân";
                ws4.Cell(1, 14).Value = "Đối sách";
                ws4.Cell(1, 15).Value = "NCC";
                ws4.Cell(1, 16).Value = "Bộ phận";
                ws4.Cell(1, 17).Value = "Ngày hoàn thành giấy lỗi";
                ws4.Cell(1, 18).Value = "Biện pháp khắc phục";
                ws4.Cell(1, 19).Value = "Ghi chú";
                ws4.Range("A1:S1").Style.Font.Bold = true;
                ws4.Range("A1:S1").Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF2CC");
                ws4.SheetView.FreezeRows(1);
                stt = 1;
                foreach (var error in errorDatas)
                {
                    row = ws4.LastRowUsed().RowNumber() + 1;
                    ws4.Cell(row, 1).Value = stt++;
                    ws4.Cell(row, 2).Value = error.DateMonth.ToString();
                    ws4.Cell(row, 3).Value = error.OrderNo;
                    ws4.Cell(row, 4).Value = error.PartName;
                    ws4.Cell(row, 5).Value = error.QtyOrder;
                    ws4.Cell(row, 6).Value = error.QtyNG;
                    ws4.Cell(row, 7).Value = error.QtyOrder != 0 ? ((decimal)error.QtyNG / error.QtyOrder * 100).ToString("0.00") + " %" : "0 %";

                    decimal ngRate = error.QtyOrder != 0 ? ((decimal)error.QtyNG / error.QtyOrder) : 0;
                    if (ngRate > 0.05m)
                    {
                        ws4.Cell(row, 7).Style.Fill.BackgroundColor = XLColor.Red;
                    }

                    ws4.Cell(row, 8).Value = error.ErrorDetected;
                    ws4.Cell(row, 9).Value = error.ErrorType;
                    ws4.Cell(row, 10).Value = error.ErrorCause;
                    ws4.Cell(row, 11).Value = error.ErrorContent;
                    ws4.Cell(row, 11).Style.Alignment.WrapText = true;
                    ws4.Cell(row, 12).Value = error.ToleranceAssessment;
                    ws4.Cell(row, 13).Value = error.ErrorCause;
                    ws4.Cell(row, 13).Style.Alignment.WrapText = true;
                    ws4.Cell(row, 14).Value = error.Countermeasure;
                    ws4.Cell(row, 14).Style.Alignment.WrapText = true;
                    ws4.Cell(row, 15).Value = error.NCC;
                    ws4.Cell(row, 16).Value = error.Department;
                    ws4.Cell(row, 17).Value = error.ErrorCompletionDate.ToString();
                    ws4.Cell(row, 18).Value = error.RemedialMeasures;
                    ws4.Cell(row, 19).Value = error.Note;
                    ws4.Cell(row, 19).Style.Alignment.WrapText = true;
                }
                lastRow = ws4.LastRowUsed().RowNumber();
                if (lastRow > 1)
                {
                    var table = ws4.Range(1, 1, lastRow, 19).CreateTable();
                    table.Theme = XLTableTheme.None;
                }
                ws4.Columns().AdjustToContents();

                var ws5 = workbook.Worksheets.Add("Vi phạm 5S");
                ws5.Cell(1, 1).Value = "STT";
                ws5.Cell(1, 2).Value = "Nội dung lỗi 5S";
                ws5.Cell(1, 3).Value = "Tháng/Năm";
                ws5.Cell(1, 4).Value = "Số lần vi phạm";
                ws5.Range("A1:D1").Style.Font.Bold = true;
                ws5.Range("A1:D1").Style.Fill.BackgroundColor = XLColor.LightCoral;
                ws5.SheetView.FreezeRows(1);
                stt = 1;
                foreach (var v5s in violation5S)
                {
                    row = ws5.LastRowUsed().RowNumber() + 1;
                    ws5.Cell(row, 1).Value = stt++;
                    ws5.Cell(row, 2).Value = v5s.Content5S;
                    ws5.Cell(row, 3).Style.Alignment.WrapText = true;
                    ws5.Cell(row, 3).Value = v5s.DateMonth.ToString();
                    ws5.Cell(row, 4).Value = v5s.Qty;
                }
                lastRow = ws5.LastRowUsed().RowNumber();
                if (lastRow > 1)
                {
                    var table = ws5.Range(1, 1, lastRow, 4).CreateTable();
                    table.Theme = XLTableTheme.None;
                }
                ws5.Columns().AdjustToContents();

                var ws6 = workbook.Worksheets.Add("Đào tạo NVM");
                ws6.Cell(1, 1).Value = "Đánh giá sau đào tạo";
                ws6.Range("A1:C1").Merge().Style.Font.Bold = true;
                ws6.Range("A1:C1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws6.Range("A1:C1").Style.Fill.BackgroundColor = XLColor.FromHtml("#B4C6E7");
                ws6.Range("A1:C1").Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);

                ws6.Cell(2, 1).Value = "STT";
                ws6.Cell(2, 2).Value = "Kỳ đào tạo";
                ws6.Cell(2, 3).Value = "Kết quả";
                ws6.Range("A2:C2").Style.Font.Bold = true;
                ws6.Range("A2:C2").Style.Fill.BackgroundColor = XLColor.FromHtml("#D9E1F2");
                ws6.Range("A2:C2").Style.Font.FontColor = XLColor.Black;

                ws6.Cell(1, 5).Value = "Kiểm tra sau đào tạo";
                ws6.Range("E1:G1").Merge().Style.Font.Bold = true;
                ws6.Range("E1:G1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws6.Range("E1:G1").Style.Fill.BackgroundColor = XLColor.FromHtml("#B4C6E7");
                ws6.Range("E1:G1").Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);

                ws6.Cell(2, 5).Value = "STT";
                ws6.Cell(2, 6).Value = "Bài kiểm tra";
                ws6.Cell(2, 7).Value = "Tổng điểm";
                ws6.Range("E2:G2").Style.Font.Bold = true;
                ws6.Range("E2:G2").Style.Fill.BackgroundColor = XLColor.FromHtml("#D9E1F2");
                ws6.Range("E2:G2").Style.Font.FontColor = XLColor.Black;

                ws6.SheetView.FreezeRows(1);
                stt = 1;
                int maxRows = 0;
                foreach (var tr in trainingresults)
                {
                    row = ws6.LastRowUsed().RowNumber() + 1;
                    ws6.Cell(row, 1).Value = stt++;
                    ws6.Cell(row, 2).Value = tr.EvaluationPeriod;
                    ws6.Cell(row, 3).Value = tr.Score + " %";

                    decimal score;
                    if (decimal.TryParse(tr.Score.ToString(), out score))
                    {
                        if (score >= 80)
                        {
                            ws6.Cell(row, 3).Style.Fill.BackgroundColor = XLColor.Green;
                        }
                        else
                        {
                            ws6.Cell(row, 3).Style.Fill.BackgroundColor = XLColor.Red;
                        }
                    }
                    maxRows = Math.Max(maxRows, row);
                }

                stt = 1;
                row = 3;
                foreach (var eta in examTrainingAnswers)
                {
                    ws6.Cell(row, 5).Value = stt++;
                    ws6.Cell(row, 6).Value = eta.ExamName;
                    ws6.Cell(row, 7).Value = $"{eta.Answer}/{eta.Total}";
                    row++;
                    maxRows = Math.Max(maxRows, row - 1);
                }

                if (maxRows >= 2)
                {
                    ws6.Range("A1:C" + maxRows).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
                    ws6.Range("E1:G" + maxRows).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
                    ws6.Range("A2:C" + maxRows).Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
                    ws6.Range("E2:G" + maxRows).Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
                }
                ws6.Columns().AdjustToContents();

                var ws7 = workbook.Worksheets.Add("Kiểm tra định kỳ");
                ws7.Cell(1, 1).Value = "STT";
                ws7.Cell(1, 2).Value = "Bài kiểm tra";
                ws7.Cell(1, 3).Value = "Tổng điểm";
                ws7.Range("A1:C1").Style.Font.Bold = true;
                ws7.Range("A1:C1").Style.Fill.BackgroundColor = XLColor.FromHtml("#C6E0B4");
                ws7.SheetView.FreezeRows(1);
                stt = 1;
                foreach (var epa in examPeriodicAnswers)
                {
                    row = ws7.LastRowUsed().RowNumber() + 1;
                    ws7.Cell(row, 1).Value = stt++;
                    ws7.Cell(row, 2).Value = epa.ExamName;
                    ws7.Cell(row, 3).Value = $"{epa.Answer}/{epa.Total}";
                }
                lastRow = ws7.LastRowUsed().RowNumber();
                if (lastRow > 1)
                {
                    var table = ws7.Range(1, 1, lastRow, 3).CreateTable();
                    table.Theme = XLTableTheme.None;
                }
                ws7.Columns().AdjustToContents();

                var ws8 = workbook.Worksheets.Add("Kiểm tra chạy thử");
                ws8.Cell(1, 1).Value = "Chạy thử lý thuyết";
                ws8.Range("A1:G1").Merge().Style.Font.Bold = true;
                ws8.Range("A1:G1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws8.Range("A1:G1").Style.Fill.BackgroundColor = XLColor.FromHtml("#F8CBAD");
                ws8.Range("A1:G1").Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);

                ws8.Cell(2, 1).Value = "STT";
                ws8.Cell(2, 2).Value = "Bài kiểm tra";
                ws8.Cell(2, 3).Value = "Đúng";
                ws8.Cell(2, 4).Value = "Sai";
                ws8.Cell(2, 5).Value = "Liệt";
                ws8.Range("A2:G2").Style.Font.Bold = true;
                ws8.Range("A2:G2").Style.Fill.BackgroundColor = XLColor.FromHtml("#FFE6CC");
                ws8.Range("A2:G2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws8.Range("A2:G2").Style.Font.FontColor = XLColor.Black;

                ws8.Cell(1, 9).Value = "Chạy thử thực hành";
                ws8.Range("I1:M1").Merge().Style.Font.Bold = true;
                ws8.Range("I1:M1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws8.Range("I1:M1").Style.Fill.BackgroundColor = XLColor.FromHtml("#F8CBAD");
                ws8.Range("I1:M1").Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);

                ws8.Cell(2, 9).Value = "STT";
                ws8.Cell(2, 10).Value = "Kỳ chạy thử";
                ws8.Cell(2, 11).Value = "Cấp độ";
                ws8.Cell(2, 12).Value = "Tên chi tiết";
                ws8.Cell(2, 13).Value = "Kết quả";
                ws8.Range("I2:M2").Style.Font.Bold = true;
                ws8.Range("I2:M2").Style.Fill.BackgroundColor = XLColor.FromHtml("#FFE6CC");
                ws8.Range("I2:M2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws8.Range("I2:M2").Style.Font.FontColor = XLColor.Black;

                ws8.SheetView.FreezeRows(1);
                int theoryLastRow = 2;

                stt = 1;
                foreach (var etra in examTrialRunAnswers)
                {
                    row = ws8.LastRowUsed().RowNumber() + 1;
                    ws8.Cell(row, 1).Value = stt++;
                    ws8.Cell(row, 2).Value = etra.ExamName;
                    ws8.Cell(row, 3).Value = etra.Correct;
                    ws8.Cell(row, 4).Value = etra.Incorrect;
                    ws8.Cell(row, 5).Value = etra.CriticalFail;

                    if (etra.CriticalFail > 0)
                    {
                        ws8.Cell(row, 5).Style.Fill.BackgroundColor = XLColor.Red;
                    }
                    theoryLastRow = row;
                }

                stt = 1;
                row = 3;
                int practiceLastRow = 2;

                foreach (var tp in testPractices)
                {
                    ws8.Cell(row, 9).Value = stt++;
                    ws8.Cell(row, 10).Value = tp.TestName;
                    ws8.Cell(row, 11).Value = tp.TestLevel;
                    ws8.Cell(row, 12).Value = tp.PartName;
                    ws8.Cell(row, 13).Value = tp.Result;

                    if (tp.Result != null && tp.Result.Contains("OK"))
                    {
                        ws8.Cell(row, 13).Style.Fill.BackgroundColor = XLColor.Green;
                    }
                    else if (tp.Result != null && tp.Result.Contains("NG"))
                    {
                        ws8.Cell(row, 13).Style.Fill.BackgroundColor = XLColor.Red;
                    }
                    practiceLastRow = row;
                    row++;
                }

                int maxDataRow = Math.Max(theoryLastRow, practiceLastRow);

                if (maxDataRow >= 2)
                {
                    ws8.Range("A1:G" + maxDataRow).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
                    ws8.Range("I1:M" + maxDataRow).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
                    ws8.Range("A2:G" + maxDataRow).Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
                    ws8.Range("I2:M" + maxDataRow).Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
                }
                ws8.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        employeeName.Replace(" ", "_") + ".xlsx"
                    );
                }
            }
        }
    }
}
