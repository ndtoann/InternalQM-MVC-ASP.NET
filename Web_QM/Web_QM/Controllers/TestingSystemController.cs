using DocumentFormat.OpenXml.Office2010.Excel;
using iText.Kernel.Pdf;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Reflection.PortableExecutable;
using Web_QM.Models;
using Web_QM.Models.ViewModels;

namespace Web_QM.Controllers
{
    [Authorize]
    public class TestingSystemController : Controller
    {
        private readonly QMContext _context;

        private readonly IWebHostEnvironment _env;

        public TestingSystemController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        public async Task<IActionResult> Index()
        {
            return View();
        }

        public async Task<IActionResult> ListExam(string department)
        {
            ViewBag.Department = department;
            var listExam = await _context.ExamPeriodics.AsNoTracking().Where(e => e.IsActive == 1 && e.ExamName.Contains(department)).OrderByDescending(e => e.Id).ToListAsync();
            return View(listExam);
        }

        public async Task<IActionResult> CNCView()
        {
            var listExamPeriodic = await _context.ExamPeriodics.AsNoTracking().Where(e => e.IsActive == 1 && e.ExamName.Contains("CNC")).OrderByDescending(e => e.Id).ToListAsync();
            var listExamTrialRun = await _context.ExamTrialRuns.AsNoTracking().Where(e => e.IsActive == 1 && e.ExamName.Contains("CNC")).OrderByDescending(e => e.Id).ToListAsync();

            ViewBag.Periodic = listExamPeriodic;
            ViewBag.TrialRun = listExamTrialRun;
            return View();
        }

        public async Task<IActionResult> TestTrainingView()
        {
            var listExam = await _context.ExamTrainings.AsNoTracking().Where(e => e.IsActive == 1).OrderByDescending(e => e.Id).ToListAsync();
            return View(listExam);
        }

        public IActionResult Search()
        {
            return View();
        }

        public async Task<IActionResult> Searching(string key)
        {
            var res = await _context.ExamPeriodicAnswers
                            .Where(answer => answer.IsShow == 1 && answer.EmployeeCode == key)
                            .Join(
                                _context.ExamPeriodics,
                                answer => answer.ExamId,
                                exam => exam.Id,
                                (answer, exam) => new { answer, exam }
                            )
                            .Join(
                                _context.Questions,
                                combined => combined.exam.Id,
                                question => question.ExamId,
                                (combined, question) => new { combined.answer, combined.exam, question }
                            )
                            .GroupBy(x => new
                            {
                                x.answer.Id,
                                x.answer.EmployeeCode,
                                x.answer.EmployeeName,
                                x.answer.TnPoint,
                                x.answer.TlPoint,
                                x.exam.ExamName,
                                x.exam.TlTotal
                            })
                            .Select(g => new SearchView
                            {
                                Id = g.Key.Id,
                                EmployeeCode = g.Key.EmployeeCode,
                                EmployeeName = g.Key.EmployeeName,
                                TnPoint = g.Key.TnPoint ?? 0,
                                TlPoint = g.Key.TlPoint ?? 0,
                                ExamName = g.Key.ExamName,
                                TnTotal = g.Count(),
                                TlTotal = g.Key.TlTotal
                            })
                            .OrderByDescending(s => s.Id)
                            .Take(5)
                            .ToListAsync();

            ViewBag.Key = key;
            return View("Search", res);
        }

        public async Task<IActionResult> SearchTrialRun()
        {
            return View();
        }

        public async Task<IActionResult> SearchingTrialRun(string key)
        {
            var employee = await _context.Employees.FirstOrDefaultAsync(e => e.EmployeeCode == key);
            if (employee == null)
            {
                return View("SearchTrialRun");
            }
            var theories = from exam in _context.ExamTrialRuns
                          join answer in _context.ExamTrialRunAnswers
                          on exam.Id equals answer.ExamTrialRunId
                          where answer.IsShow == 1 && answer.EmployeeCode == key
                          select new
                          {
                              Id = answer.Id,
                              ExamName = exam.ExamName,
                              TestLevel = exam.TestLevel,
                              Correct = answer.MultipleChoiceCorrect + answer.EssayCorrect,
                              InCorrect = answer.MultipleChoiceInCorrect + answer.EssayInCorrect,
                              CriticalFail = answer.MultipleChoiceFail + answer.EssayFail,
                              EmployeeName = answer.EmployeeName,
                              EmployeeCode = answer.EmployeeCode,
                          };
            ViewBag.Theories = await theories.OrderByDescending(o => o.Id).Take(1).ToListAsync();

            var practices = await _context.TestPractices.OrderByDescending(o => o.Id).FirstOrDefaultAsync(p => p.EmployeeId == employee.Id);
            ViewBag.Practices = practices;

            if(practices != null)
            {
                var practiceDetail = await _context.TestPracticeDetails.Where(p => p.TestPracticeId == practices.Id).ToListAsync();
                ViewBag.PracticeDetail = practiceDetail;
            }

            ViewBag.Key = key;
            return View("SearchTrialRun");
        }

        public async Task<IActionResult> Prepare(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var res = await _context.ExamPeriodics.AsNoTracking().Where(q => q.IsActive == 1).FirstOrDefaultAsync(x => x.Id == id);
            if (res == null)
            {
                return NotFound();
            }
            return View(res);
        }

        public async Task<IActionResult> CNCPrepare(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var res = await _context.ExamTrialRuns.AsNoTracking().Where(q => q.IsActive == 1).FirstOrDefaultAsync(x => x.Id == id);
            if (res == null)
            {
                return NotFound();
            }
            return View(res);
        }

        public async Task<IActionResult> TestTrainingPrepare(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var res = await _context.ExamTrainings.AsNoTracking().Where(q => q.IsActive == 1).FirstOrDefaultAsync(x => x.Id == id);
            if (res == null)
            {
                return NotFound();
            }
            return View(res);
        }

        public async Task<IActionResult> Start(long id, string emplCode)
        {
            if (id == null)
            {
                return NotFound();
            }
            if (emplCode == null)
            {
                TempData["ErrorStart"] = "Mã nhân viên không hợp lệ!";
                return RedirectToAction(nameof(Prepare), new { id = id });
            }

            var exam = await _context.ExamPeriodics.Where(e => e.IsActive == 1).FirstOrDefaultAsync(x => x.Id == id);
            if (exam == null)
            {
                return NotFound();
            }

            var empl = await _context.Employees.FirstOrDefaultAsync(e => e.EmployeeCode == emplCode);
            if (empl == null)
            {
                TempData["ErrorStart"] = "Mã nhân viên không hợp lệ!";
                return RedirectToAction(nameof(Prepare), new { id = id });
            }

            var emplDouble = await _context.ExamPeriodicAnswers.CountAsync(e => e.EmployeeCode == emplCode && e.ExamId == id);
            if (emplDouble > 0)
            {
                TempData["ErrorStart"] = "Bạn đã làm bài kiểm tra này rồi!";
                return RedirectToAction(nameof(Prepare), new { id = id });
            }

            string pdfFileName = exam.EssayQuestion;
            if (string.IsNullOrEmpty(pdfFileName))
            {
                pdfFileName = "";
            }
            string pdfPath = Path.Combine(_env.WebRootPath, "files", "essay_questions", pdfFileName);
            int pageCount = 0;
            if (System.IO.File.Exists(pdfPath))
            {
                pageCount = GetPdfPageCount(pdfPath);
            }
            ViewBag.PdfPageCount = pageCount;
            ViewBag.EmplCode = emplCode;
            ViewBag.EmplName = empl.EmployeeName;
            ViewBag.EmplDepartment = empl.Department;

            var listQuestion = await _context.Questions.AsNoTracking().Where(q => q.ExamId == id).OrderBy(o => o.DisplayOrder).ToListAsync();
            var viewData = new ExamDetailView
            {
                exam = exam,
                questions = listQuestion
            };
            return View(viewData);
        }

        public async Task<IActionResult> CNCStart(long id, string emplCode)
        {
            if (id == null)
            {
                return NotFound();
            }
            if (emplCode == null)
            {
                TempData["ErrorStart"] = "Mã nhân viên không hợp lệ!";
                return RedirectToAction(nameof(CNCPrepare), new { id = id });
            }

            var exam = await _context.ExamTrialRuns.Where(e => e.IsActive == 1).FirstOrDefaultAsync(x => x.Id == id);
            if (exam == null)
            {
                return NotFound();
            }

            var empl = await _context.Employees.FirstOrDefaultAsync(e => e.EmployeeCode == emplCode);
            if (empl == null)
            {
                TempData["ErrorStart"] = "Mã nhân viên không hợp lệ!";
                return RedirectToAction(nameof(CNCPrepare), new { id = id });
            }

            var emplDouble = await _context.ExamTrialRunAnswers.CountAsync(e => e.EmployeeCode == emplCode && e.ExamTrialRunId == id);
            if (emplDouble > 0)
            {
                TempData["ErrorStart"] = "Bạn đã làm bài kiểm tra này rồi!";
                return RedirectToAction(nameof(CNCPrepare), new { id = id });
            }

            string pdfFileName = exam.EssayQuestion;
            if (string.IsNullOrEmpty(pdfFileName))
            {
                pdfFileName = "";
            }
            string pdfPath = Path.Combine(_env.WebRootPath, "files", "essay_questions", pdfFileName);
            int pageCount = 0;
            if (System.IO.File.Exists(pdfPath))
            {
                pageCount = GetPdfPageCount(pdfPath);
            }
            ViewBag.PdfPageCount = pageCount;
            ViewBag.EmplCode = emplCode;
            ViewBag.EmplName = empl.EmployeeName;
            ViewBag.EmplDepartment = empl.Department;

            var listQuestion = await _context.QuestionTrialRuns.AsNoTracking().Where(q => q.ExamTrialRunId == id).OrderBy(o => o.DisplayOrder).ToListAsync();
            var viewData = new ExamTrialRunDetailView
            {
                exam = exam,
                questions = listQuestion
            };
            return View(viewData);
        }

        public async Task<IActionResult> TestTrainingStart(long id, string emplCode)
        {
            if (id == null)
            {
                return NotFound();
            }
            if (emplCode == null)
            {
                TempData["ErrorStart"] = "Mã nhân viên không hợp lệ!";
                return RedirectToAction(nameof(TestTrainingPrepare), new { id = id });
            }

            var exam = await _context.ExamTrainings.Where(e => e.IsActive == 1).FirstOrDefaultAsync(x => x.Id == id);
            if (exam == null)
            {
                return NotFound();
            }

            var empl = await _context.Employees.FirstOrDefaultAsync(e => e.EmployeeCode == emplCode);
            if (empl == null)
            {
                TempData["ErrorStart"] = "Mã nhân viên không hợp lệ!";
                return RedirectToAction(nameof(TestTrainingPrepare), new { id = id });
            }

            var emplDouble = await _context.ExamTrainingAnswers.CountAsync(e => e.EmployeeCode == emplCode && e.ExamTrainingId == id);
            if (emplDouble > 0)
            {
                TempData["ErrorStart"] = "Bạn đã làm bài kiểm tra này rồi!";
                return RedirectToAction(nameof(TestTrainingPrepare), new { id = id });
            }

            string pdfFileName = exam.EssayQuestion;
            if (string.IsNullOrEmpty(pdfFileName))
            {
                pdfFileName = "";
            }
            string pdfPath = Path.Combine(_env.WebRootPath, "files", "essay_questions", pdfFileName);
            int pageCount = 0;
            if (System.IO.File.Exists(pdfPath))
            {
                pageCount = GetPdfPageCount(pdfPath);
            }
            ViewBag.PdfPageCount = pageCount;
            ViewBag.EmplCode = emplCode;
            ViewBag.EmplName = empl.EmployeeName;
            ViewBag.EmplDepartment = empl.Department;

            var listQuestion = await _context.QuestionTrainings.AsNoTracking().Where(q => q.ExamTrainingId == id).OrderBy(o => o.DisplayOrder).ToListAsync();
            var viewData = new ExamTrainingDetailView
            {
                exam = exam,
                questions = listQuestion
            };
            return View(viewData);
        }

        [HttpPost]
        public async Task<IActionResult> SaveAnswer([FromBody] ExamPeriodicAnswer employeeAnswer)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(new { status = false, message = "Thông tin không hợp lệ."});
            }

            var examID = employeeAnswer.ExamId;
            var emplName = employeeAnswer.EmployeeName;
            var emplCode = employeeAnswer.EmployeeCode;

            var checkEmpl = await _context.Employees.CountAsync(e => e.EmployeeCode == emplCode);
            if(checkEmpl == 0)
            {
                return BadRequest(new { status = false, message = "Mã nhân viên không hợp lệ." });
            }

            var emplDouble = await _context.ExamPeriodicAnswers.CountAsync(e => e.EmployeeCode == emplCode && e.ExamId == examID);
            if (emplDouble > 0)
            {
                return BadRequest(new { status = false, message = "Bạn đã làm bài kiểm tra này rồi." });
            }

            var listAnswer = employeeAnswer.ListAnswer;

            employeeAnswer.TnPoint = CalculateTnPoint(employeeAnswer.ListAnswer, employeeAnswer.ExamId).Result;

            employeeAnswer.IsShow = 0;
            employeeAnswer.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
            try
            {
                _context.ExamPeriodicAnswers.Add(employeeAnswer);
                await _context.SaveChangesAsync();

                var res = new { status = true, message = "Lưu bài làm thành công." };
                return Ok(res);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi lưu bài làm: {ex.Message}");

                var res = new { status = false, message = "Lỗi server: " + ex.Message };
                return StatusCode(500, res);
            }
        }

        private async Task<int> CalculateTnPoint(string listAnswer, long? examId)
        {
            var point = 0;

            if (string.IsNullOrEmpty(listAnswer))
            {
                return 0;
            }
            var employeeAnswers = new Dictionary<int, string>();
            foreach (var pair in listAnswer.Split('-'))
            {
                var parts = pair.Split('.');
                if (parts.Length == 2 && int.TryParse(parts[0], out int questionIndex))
                {
                    employeeAnswers[questionIndex] = parts[1];
                }
            }

            var allCorrectAnswersFromDb = await _context.Questions
                .Where(q => q.ExamId == examId)
                .OrderBy(q => q.DisplayOrder)
                .ToListAsync();

            var correctAnswersIndexed = new Dictionary<int, string>();
            for (int i = 0; i < allCorrectAnswersFromDb.Count; i++)
            {
                correctAnswersIndexed[i + 1] = allCorrectAnswersFromDb[i].CorrectOption;
            }

            foreach (var userAnswer in employeeAnswers)
            {
                int questionIndex = userAnswer.Key;
                string selectedOption = userAnswer.Value;

                if (correctAnswersIndexed.ContainsKey(questionIndex))
                {
                    if (selectedOption.Trim().Equals(correctAnswersIndexed[questionIndex].Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        point++;
                    }
                }
            }

            return point;
        }

        [HttpPost]
        public async Task<IActionResult> SaveTrainingAnswer([FromBody] ExamTrainingAnswer employeeAnswer)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(new { status = false, message = "Thông tin không hợp lệ." });
            }

            var examID = employeeAnswer.ExamTrainingId;
            var emplName = employeeAnswer.EmployeeName;
            var emplCode = employeeAnswer.EmployeeCode;

            var checkEmpl = await _context.Employees.CountAsync(e => e.EmployeeCode == emplCode);
            if (checkEmpl == 0)
            {
                return BadRequest(new { status = false, message = "Mã nhân viên không hợp lệ." });
            }

            var emplDouble = await _context.ExamPeriodicAnswers.CountAsync(e => e.EmployeeCode == emplCode && e.ExamId == examID);
            if (emplDouble > 0)
            {
                return BadRequest(new { status = false, message = "Bạn đã làm bài kiểm tra này rồi." });
            }

            var listAnswer = employeeAnswer.ListAnswer;

            employeeAnswer.TnPoint = CalculateTnPointTraining(employeeAnswer.ListAnswer, employeeAnswer.ExamTrainingId).Result;

            employeeAnswer.IsShow = 0;
            employeeAnswer.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
            try
            {
                _context.ExamTrainingAnswers.Add(employeeAnswer);
                await _context.SaveChangesAsync();

                var res = new { status = true, message = "Lưu bài làm thành công." };
                return Ok(res);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi lưu bài làm: {ex.Message}");

                var res = new { status = false, message = "Lỗi server: " + ex.Message };
                return StatusCode(500, res);
            }
        }

        private async Task<int> CalculateTnPointTraining(string listAnswer, long? examId)
        {
            var point = 0;

            if (string.IsNullOrEmpty(listAnswer))
            {
                return 0;
            }

            var employeeAnswers = new Dictionary<int, string>();
            foreach (var pair in listAnswer.Split('-'))
            {
                var parts = pair.Split('.');
                if (parts.Length == 2 && int.TryParse(parts[0], out int questionIndex))
                {
                    employeeAnswers[questionIndex] = parts[1];
                }
            }

            var allCorrectAnswersFromDb = await _context.QuestionTrainings
                .Where(q => q.ExamTrainingId == examId)
                .OrderBy(q => q.DisplayOrder)
                .ToListAsync();

            var correctAnswersIndexed = new Dictionary<int, string>();
            for (int i = 0; i < allCorrectAnswersFromDb.Count; i++)
            {
                correctAnswersIndexed[i + 1] = allCorrectAnswersFromDb[i].CorrectOption;
            }

            foreach (var userAnswer in employeeAnswers)
            {
                int questionIndex = userAnswer.Key;
                string selectedOption = userAnswer.Value;

                if (correctAnswersIndexed.ContainsKey(questionIndex))
                {
                    if (selectedOption.Trim().Equals(correctAnswersIndexed[questionIndex].Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        point++;
                    }
                }
            }

            return point;
        }

        [HttpPost]
        public async Task<IActionResult> SaveTrialRunAnswer([FromBody] ExamTrialRunAnswer employeeAnswer)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(new { status = false, message = "Thông tin không hợp lệ." });
            }
            var examID = employeeAnswer.ExamTrialRunId;
            var emplName = employeeAnswer.EmployeeName;
            var emplCode = employeeAnswer.EmployeeCode;

            var checkEmpl = await _context.Employees.CountAsync(e => e.EmployeeCode == emplCode);
            if (checkEmpl == 0)
            {
                return BadRequest(new { status = false, message = "Mã nhân viên không hợp lệ." });
            }

            var emplDouble = await _context.ExamTrialRunAnswers.CountAsync(e => e.EmployeeCode == emplCode && e.ExamTrialRunId == examID);
            if (emplDouble > 0)
            {
                return BadRequest(new { status = false, message = "Bạn đã làm bài kiểm tra này rồi." });
            }

            var listAnswer = employeeAnswer.ListAnswer;

            var eS = await CalculateExamScore(employeeAnswer.ListAnswer, employeeAnswer.ExamTrialRunId);
            employeeAnswer.MultipleChoiceCorrect = eS.Correct;
            employeeAnswer.MultipleChoiceInCorrect = eS.Incorrect;
            employeeAnswer.MultipleChoiceFail = eS.CriticalFail;

            employeeAnswer.IsShow = 0;
            employeeAnswer.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
            try
            {
                _context.ExamTrialRunAnswers.Add(employeeAnswer);
                await _context.SaveChangesAsync();

                var res = new { status = true, message = "Lưu bài làm thành công." };
                return Ok(res);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi lưu bài làm: {ex.Message}");

                var res = new { status = false, message = "Lỗi server: " + ex.Message };
                return StatusCode(500, res);
            }
        }

        private async Task<ExamScore> CalculateExamScore(string listAnswer, long? examId)
        {
            var score = new ExamScore();

            var employeeAnswers = new Dictionary<int, string>();
            if (!string.IsNullOrEmpty(listAnswer))
            {
                foreach (var pair in listAnswer.Split('-'))
                {
                    var parts = pair.Split('.');
                    if (parts.Length == 2 && int.TryParse(parts[0], out int questionIndex))
                    {
                        employeeAnswers[questionIndex] = parts[1];
                    }
                }
            }

            var allQuestionsFromDb = await _context.QuestionTrialRuns
                                                    .Where(q => q.ExamTrialRunId == examId)
                                                    .OrderBy(q => q.DisplayOrder)
                                                    .ToListAsync();

            for (int i = 0; i < allQuestionsFromDb.Count; i++)
            {
                var question = allQuestionsFromDb[i];
                int questionIndex = i + 1;

                string employeeAnswer;

                if (employeeAnswers.TryGetValue(questionIndex, out employeeAnswer))
                {
                    if (employeeAnswer.Trim().Equals(question.CorrectOption.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        score.Correct++;
                    }
                    else
                    {
                        if (question.IsCritical == 1)
                        {
                            score.CriticalFail++;
                        }
                        else
                        {
                            score.Incorrect++;
                        }
                    }
                }
                else
                {
                    if (question.IsCritical == 1)
                    {
                        score.CriticalFail++;
                    }
                    else
                    {
                        score.Incorrect++;
                    }
                }
            }

            return score;
        }

        public int GetPdfPageCount(string filePath)
        {
            if (!System.IO.File.Exists(filePath))
            {
                return 0;
            }
            int pageCount = 0;
            try
            {
                using (var reader = new PdfReader(filePath))
                {
                    using (var pdfDocument = new PdfDocument(reader))
                    {
                        pageCount = pdfDocument.GetNumberOfPages();
                    }
                }
            }
            catch (Exception ex)
            {
                return 0;
            }

            return pageCount;
        }
    }
}
