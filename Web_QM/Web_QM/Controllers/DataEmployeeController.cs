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

        public DataEmployeeController( QMContext context)
        {
            _context = context;
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

            var userDepartment = User.FindFirst("Department")?.Value;

            var query = _context.Employees.AsNoTracking();

            if (!canViewAll)
            {
                if (string.IsNullOrEmpty(userDepartment))
                {
                    return Json(new { data = new List<object>(), message = "Không có quyền truy cập dữ liệu nhân viên!" });
                }
                query = query.Where(e => e.Department == userDepartment);
            }

            if (!string.IsNullOrEmpty(search))
            {
                var lowerSearch = search.Trim().ToLower();
                query = query.Where(e =>
                    e.EmployeeName.ToLower().Contains(lowerSearch) ||
                    e.EmployeeCode.ToLower().Contains(lowerSearch)
                );
            }

            if (canViewAll && !string.IsNullOrEmpty(department))
            {
                query = query.Where(e => e.Department == department);
            }

            query = query.OrderBy(o => o.EmployeeCode);

            var employees = await query.ToListAsync();

            return Json(new { data = employees });
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

            // xử lý phản hồi từ tổ trưởng
            var feedbacks = await _context.Feedbacks.AsNoTracking().Where(f => f.EmployeeId == id).OrderByDescending(f => f.Id).Take(50).ToListAsync();

            ViewBag.Feedbacks = feedbacks;

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

            // xử lý kết quả kiểm tra của nhân viên
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
                    Note = joined.EmployeeAnswer.Note,
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
                feedback.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");

                _context.Feedbacks.Add(feedback);

                var empl = await _context.Employees.FirstOrDefaultAsync(e => e.Id == feedback.EmployeeId);

                Notification nt = new Notification()
                {
                    Message = "Nhân viên " + empl.EmployeeCode + "-" + empl.EmployeeName + " có phản hồi mới",
                    IsRead = 0,
                    CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy")
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
    }
}
