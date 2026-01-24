using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Globalization;
using Web_QM.Models;
using Web_QM.Models.ViewModels;

namespace Web_QM.Controllers
{
    [Authorize]
    public class ProfileController : Controller
    {
        private readonly QMContext _context;

        public ProfileController(QMContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> Index()
        {
            var employeeCodeIsLogin = User.FindFirst("EmployeeCode")?.Value;
            if (employeeCodeIsLogin == null)
            {
                return NotFound();
            }

            var employee = await _context.Employees
                    .AsNoTracking()
                    .Where(e => e.EmployeeCode == employeeCodeIsLogin)
                    .FirstOrDefaultAsync();

            if (employee == null)
            {
                return NotFound();
            }

            // xử lý dữ liệu đào tạo
            var evaluationPeriodsWithScore = await _context.EmployeeTrainingResults
                                                 .AsNoTracking()
                                                 .Where(tr => tr.EmployeeId == employee.Id)
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
            var testPractices = await _context.TestPractices.AsNoTracking().Where(t => t.EmployeeId == employee.Id).OrderByDescending(o => o.Id).Take(100).ToListAsync();
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
                .Select(v => new
                {
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

        public async Task<IActionResult> GetFeedbacks(long employeeId)
        {
            if (employeeId == 0)
            {
                return BadRequest();
            }
            var empl = await _context.Employees.FirstOrDefaultAsync(e => e.Id == employeeId);
            if (empl == null)
            {
                return BadRequest();
            }
            IQueryable<Feedback> query = _context.Feedbacks.Where(f => f.EmployeeId == employeeId && f.Status == 1);
            var feedbacks = await query.OrderByDescending(f => f.Id).ToListAsync();
            return Json(feedbacks);
        }

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

        public async Task<IActionResult> GetTrainings(string type)
        {
            var trainings = await _context.Trainings.Where(t => t.Type == type).ToListAsync();
            return Json(trainings);
        }

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


        public async Task<IActionResult> Timesheet()
        {
            return View();
        }

        public async Task<IActionResult> Timekeeping(long timesheetId)
        {
            var timesheet = await _context.Timesheets
                .FirstOrDefaultAsync(t => t.Id == timesheetId && t.EmployeeId == GetCurrentEmployeeId());

            if (timesheet == null) return NotFound();

            if (timesheet.DateMonth.Length != 7 || !DateTime.TryParseExact(timesheet.DateMonth + "-01", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime firstDayOfMonth))
            {
                return BadRequest("Định dạng DateMonth không hợp lệ.");
            }

            ViewData["Timesheet"] = timesheet;
            ViewData["MonthYear"] = firstDayOfMonth.ToString("MM/yyyy");

            return View(timesheet);
        }

        public async Task<IActionResult> GetTimesheets(string dateMonth)
        {
            IQueryable<Timesheet> query = _context.Timesheets
                .Where(t => t.EmployeeId == GetCurrentEmployeeId());

            if (!string.IsNullOrEmpty(dateMonth))
            {
                query = query.Where(t => t.DateMonth == dateMonth);
            }

            var data = await query
                .OrderByDescending(t => t.DateMonth)
                .Select(t => new { t.Id, t.DateMonth, t.Status })
                .ToListAsync();

            return Json(data);
        }

        [HttpPost]
        public async Task<IActionResult> AddTimesheet([FromBody] Timesheet model)
        {
            if (model == null || model.DateMonth.Length != 7 || !DateTime.TryParseExact(model.DateMonth + "-01", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            {
                return BadRequest(new { success = false, message = "Định dạng tháng không hợp lệ." });
            }

            var existing = await _context.Timesheets
                .FirstOrDefaultAsync(t => t.EmployeeId == GetCurrentEmployeeId() && t.DateMonth == model.DateMonth);

            if (existing != null)
            {
                return BadRequest(new { success = false, message = $"Bảng chấm công tháng {model.DateMonth} đã tồn tại." });
            }

            model.EmployeeId = GetCurrentEmployeeId();
            model.Status = 1;

            _context.Timesheets.Add(model);
            await _context.SaveChangesAsync();

            return Json(new { success = true, record = model });
        }

        [HttpPost]
        public async Task<IActionResult> ApproveTimesheet([FromBody] Timesheet model)
        {
            var sheet = await _context.Timesheets.FindAsync(model.Id);

            if (sheet == null)
            {
                return NotFound(new { success = false, message = "Phiếu chấm công không tồn tại." });
            }

            if (sheet.Status == 2)
            {
                return BadRequest(new { success = false, message = "Phiếu này đã được duyệt rồi." });
            }

            sheet.Status = 2;

            _context.Timesheets.Update(sheet);
            await _context.SaveChangesAsync();

            return Json(new { success = true, newStatus = 2 });
        }

        public async Task<IActionResult> DeleteTimesheet([FromBody] Timesheet model)
        {
            var sheet = await _context.Timesheets.FindAsync(model.Id);

            if (sheet == null)
            {
                return NotFound(new { success = false, message = "Phiếu chấm công không tồn tại." });
            }
            if (sheet.Status == 2 || sheet.Status == 3)
            {
                return BadRequest(new { success = false, message = "Phiếu này đã được duyệt rồi." });
            }

            var recordsToDelete = await _context.Timekeepings
                .Where(t => t.TimesheetId == sheet.Id)
                .ToListAsync();

            _context.Timekeepings.RemoveRange(recordsToDelete);

            _context.Timesheets.Remove(sheet);
            await _context.SaveChangesAsync();

            return Json(new { success = true });
        }

        public async Task<IActionResult> GetTimekeepingData(long timesheetId)
        {
            var timesheet = await _context.Timesheets
                .FirstOrDefaultAsync(t => t.Id == timesheetId && t.EmployeeId == GetCurrentEmployeeId());

            if (timesheet == null) return NotFound();

            var data = await _context.Timekeepings
                .Where(t => t.TimesheetId == timesheetId)
                .Select(t => new
                {
                    t.Id,
                    WorkDate = t.WorkDate.ToString("yyyy-MM-dd"),
                    t.TimeIn,
                    t.TimeOut,
                    t.Shift,
                    t.TotalHours,
                    t.Note
                })
                .ToListAsync();

            return Json(new { Timesheet = timesheet, TimekeepingData = data });
        }

        [HttpPost]
        public async Task<IActionResult> SaveTimekeeping([FromBody] Timekeeping model)
        {
            if (model == null) return BadRequest();

            var timesheet = await _context.Timesheets
                .FirstOrDefaultAsync(t => t.Id == model.TimesheetId && t.EmployeeId == GetCurrentEmployeeId());

            if (timesheet == null) return Unauthorized();

            if (timesheet.Status == 2 || timesheet.Status == 3)
            {
                return BadRequest(new { success = false, message = "Phiếu đã gửi duyệt hoặc đã duyệt, không thể chỉnh sửa." });
            }

            bool isTimeValid = model.TimeIn.HasValue && model.TimeOut.HasValue;
            bool hasTimeInput = isTimeValid && (model.TimeIn.Value != TimeSpan.Zero || model.TimeOut.Value != TimeSpan.Zero);
            bool hasNoteInput = !string.IsNullOrWhiteSpace(model.Note);

            if (hasTimeInput)
            {
                if (model.TimeIn == model.TimeOut && (model.TimeIn.Value != TimeSpan.Zero || model.TimeOut.Value != TimeSpan.Zero))
                {
                    return BadRequest(new { success = false, message = "Giờ vào và Giờ ra không được trùng nhau." });
                }
                if (model.TimeIn >= model.TimeOut && model.Shift != "Ca 3" && model.Shift != "Ca 5")
                {
                    return BadRequest(new { success = false, message = "Giờ vào phải nhỏ hơn." });
                }
            }
            else if (model.TimeIn.HasValue && model.TimeIn.Value != TimeSpan.Zero)
            {
                return BadRequest(new { success = false, message = "Vui lòng nhập Giờ ra." });
            }
            else if (model.TimeOut.HasValue && model.TimeOut.Value != TimeSpan.Zero)
            {
                return BadRequest(new { success = false, message = "Vui lòng nhập Giờ vào." });
            }


            if (!hasTimeInput && !hasNoteInput)
            {
                return BadRequest(new { success = false, message = "Vui lòng nhập giờ Vào/Ra hoặc ghi chú." });
            }

            model.TotalHours = CalculateTotalHours(model);

            var existingRecord = await _context.Timekeepings
                .FirstOrDefaultAsync(t => t.TimesheetId == model.TimesheetId && t.WorkDate == model.WorkDate);

            if (existingRecord == null)
            {
                model.Id = 0;
                _context.Timekeepings.Add(model);
            }
            else
            {
                existingRecord.TimeIn = model.TimeIn;
                existingRecord.TimeOut = model.TimeOut;
                existingRecord.Shift = model.Shift;
                existingRecord.TotalHours = model.TotalHours;
                existingRecord.Note = model.Note;

                _context.Timekeepings.Update(existingRecord);
                model.Id = existingRecord.Id;
            }

            await _context.SaveChangesAsync();

            var returnRecord = new
            {
                model.Id,
                WorkDate = model.WorkDate.ToString("yyyy-MM-dd"),
                model.TimeIn,
                model.TimeOut,
                model.Shift,
                model.TotalHours,
                model.Note
            };

            return Json(new { success = true, record = returnRecord });
        }

        public async Task<IActionResult> DeleteTimekeeping([FromBody] Timekeeping model)
        {
            var sheet = await _context.Timesheets
                .FirstOrDefaultAsync(t => t.Id == model.TimesheetId && t.EmployeeId == GetCurrentEmployeeId());

            if (sheet == null) return Unauthorized();

            if (sheet.Status == 2 || sheet.Status == 3)
            {
                return BadRequest(new { success = false, message = "Phiếu đã gửi duyệt hoặc đã duyệt, không thể xóa." });
            }

            var recordToDelete = await _context.Timekeepings
                .FirstOrDefaultAsync(t => t.TimesheetId == model.TimesheetId && t.WorkDate == model.WorkDate);

            if (recordToDelete == null)
            {
                return NotFound(new { message = "Không tìm thấy dữ liệu để xóa." });
            }

            _context.Timekeepings.Remove(recordToDelete);
            await _context.SaveChangesAsync();
            return Json(new { success = true });
        }

        private decimal CalculateTotalHours(Timekeeping model)
        {
            if (!model.TimeIn.HasValue || !model.TimeOut.HasValue || model.Shift == null)
            {
                return 0.0m;
            }

            TimeSpan timeIn = model.TimeIn.Value;
            TimeSpan timeOut = model.TimeOut.Value;
            string shift = model.Shift;

            var breakTimes = new Dictionary<string, (TimeSpan start, TimeSpan end, bool overnight)>
            {
                { "Hành chính", (new TimeSpan(12, 0, 0), new TimeSpan(13, 0, 0), false) },
                { "Ca 1", (new TimeSpan(12, 0, 0), new TimeSpan(13, 0, 0), false) },
                { "Ca 2", (new TimeSpan(18, 0, 0), new TimeSpan(19, 0, 0), false) },
                { "Ca 4", (new TimeSpan(12, 0, 0), new TimeSpan(13, 0, 0), false) },
                { "Ca 3", (new TimeSpan(1, 0, 0), new TimeSpan(2, 0, 0), true) },
                { "Ca 5", (new TimeSpan(1, 0, 0), new TimeSpan(2, 0, 0), true) }
            };

            double totalDurationInHours;

            if (timeOut > timeIn)
            {
                totalDurationInHours = (timeOut - timeIn).TotalHours;
            }
            else if (timeOut < timeIn)
            {
                totalDurationInHours = (timeOut - timeIn).TotalHours + 24;
            }
            else
            {
                return 0.0m;
            }

            decimal totalHours = (decimal)totalDurationInHours;

            if (shift != null && breakTimes.ContainsKey(shift))
            {
                var (breakStart, breakEnd, isOvernight) = breakTimes[shift];
                decimal deductedHours = 0.0m;

                if (isOvernight)
                {
                    TimeSpan breakEndNight = breakEnd.Add(TimeSpan.FromHours(24));
                    TimeSpan timeOutNight = timeOut.Add(TimeSpan.FromHours(24));

                    if (timeOutNight >= breakEndNight && totalHours > 1.0m)
                    {
                        deductedHours = 1.0m;
                    }
                }
                else
                {
                    TimeSpan overlapStart = timeIn > breakStart ? timeIn : breakStart;
                    TimeSpan overlapEnd = timeOut < breakEnd ? timeOut : breakEnd;

                    if (overlapEnd > overlapStart)
                    {
                        double overlapHours = (overlapEnd - overlapStart).TotalHours;
                        deductedHours = (decimal)Math.Min(overlapHours, 1.0);
                    }
                }

                totalHours -= deductedHours;
            }

            return Math.Max(0.0m, totalHours);
        }

        private long GetCurrentEmployeeId()
        {
            return long.Parse(User.Claims.FirstOrDefault(c => c.Type == "EmployeeId")?.Value ?? "0");
        }
    }
}
