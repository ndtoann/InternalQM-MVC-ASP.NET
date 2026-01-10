using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models.ViewModels;
using Web_QM.Models;

namespace Web_QM.Areas.HR.Controllers
{
    [Area("HR")]
    public class ResultTrialRunController : Controller
    {
        private readonly QMContext _context;

        public ResultTrialRunController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ViewResultExam")]
        public async Task<IActionResult> Index(string name)
        {
            var examStats = await _context.ExamTrialRuns
                                    .GroupJoin(
                                        _context.ExamTrialRunAnswers,
                                        exam => exam.Id,
                                        answer => answer.ExamTrialRunId,
                                        (exam, employeeAnswers) => new ExamTrialRunAnswerView
                                        {
                                            examId = exam.Id,
                                            examName = exam.ExamName,
                                            level = exam.TestLevel,
                                            employeeCount = employeeAnswers.Count()
                                        }
                                    )
                                    .Where(x => string.IsNullOrEmpty(name) || x.examName.ToLower().Contains(name.ToLower()))
                                    .OrderByDescending(e => e.examId)
                                    .Take(200)
                                    .ToListAsync();

            ViewBag.Name = name;
            return View(examStats);
        }

        [Authorize(Policy = "ViewResultExam")]
        public async Task<IActionResult> Detail(long examid)
        {
            var exam = await _context.ExamTrialRuns.AsNoTracking().FirstOrDefaultAsync(e => e.Id == examid);

            if (exam == null)
            {
                return NotFound();
            }

            var listEmpl = await _context.ExamTrialRunAnswers
                .AsNoTracking()
                .Where(e => e.ExamTrialRunId == examid)
                .OrderBy(e => e.EmployeeCode)
                .ToListAsync();

            var questionsFromDb = await _context.QuestionTrialRuns
                .AsNoTracking()
                .Where(q => q.ExamTrialRunId == examid)
                .OrderBy(q => q.DisplayOrder)
                .ToListAsync();

            var questions = questionsFromDb
                .Select((q, index) => new
                {
                    QuestionId = q.Id,
                    QuestionNumber = index + 1,
                    QuestionText = q.QuestionText,
                    OptionA = q.OptionA,
                    OptionB = q.OptionB,
                    OptionC = q.OptionC,
                    OptionD = q.OptionD,
                    CorrectOption = q.CorrectOption,
                    IsCritical = q.IsCritical
                })
                .ToList();

            ViewBag.Questions = questions;
            ViewBag.ExamId = examid;
            ViewBag.ExamName = exam?.ExamName ?? "N/A";
            ViewBag.TestLevel = exam?.TestLevel ?? "N/A";
            ViewBag.CountQuestion = questions.Count;

            return View(listEmpl);
        }

        [Authorize(Policy = "EditResultExam")]
        [HttpPost]
        public async Task<IActionResult> UpdatePoint(ExamTrialRunAnswer employeeAnswer, IFormFile pdfFile)
        {
            if (employeeAnswer == null)
            {
                return NotFound();
            }
            if (employeeAnswer.EssayCorrect < 0 || employeeAnswer.EssayInCorrect < 0 || employeeAnswer.EssayFail < 0)
            {
                return NotFound();
            }
            try
            {
                string essayResultPdfPath = await SavePdfFile(employeeAnswer.ExamTrialRunId, employeeAnswer.EmployeeCode, pdfFile);
                if (!string.IsNullOrEmpty(essayResultPdfPath))
                {
                    employeeAnswer.EssayResultPDF = essayResultPdfPath;
                }
                employeeAnswer.UpdatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.ExamTrialRunAnswers.Update(employeeAnswer);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Cập nhật thành công!";
                return RedirectToAction(nameof(Detail), new { examid = employeeAnswer.ExamTrialRunId });
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
                return RedirectToAction(nameof(Detail), new { examid = employeeAnswer.ExamTrialRunId });
            }
        }

        private async Task<string> SavePdfFile(long examId, string employeeCode, IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return null;
            }

            if (file.ContentType != "application/pdf")
            {
                return null;
            }

            var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files", "answers");
            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }
            var fileName = $"Đáp án chạy thử-{examId}-{employeeCode}.pdf";
            var filePath = Path.Combine(uploadsFolder, fileName);
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }
            return $"/files/answers/{fileName}";
        }

        [Authorize(Policy = "DeleteResultExam")]
        public async Task<IActionResult> Delete(long id)
        {
            if(id == null)
            {
                return NotFound();
            }
            var emplAnswerToDelete = await _context.ExamTrialRunAnswers.FindAsync(id);
            if (emplAnswerToDelete == null)
            {
                return NotFound();
            }

            try
            {
                DeletePdfFile(emplAnswerToDelete.ExamTrialRunId, emplAnswerToDelete.EmployeeCode);
                _context.ExamTrialRunAnswers.Remove(emplAnswerToDelete);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa bài làm nhân viên này!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Detail), new { examid = emplAnswerToDelete.ExamTrialRunId });
        }

        private async Task<bool> DeletePdfFile(long examId, string employeeCode)
        {
            try
            {
                var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files", "answers");
                var fileName = $"Đáp án chạy thử-{examId}-{employeeCode}.pdf";
                var filePath = Path.Combine(uploadsFolder, fileName);

                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        [Authorize(Policy = "ViewResultExam")]
        public async Task<IActionResult> ExportToExcel(long? examId)
        {
            if (examId == null)
            {
                return NotFound();
            }
            var exam = await _context.ExamTrialRuns.FirstOrDefaultAsync(e => e.Id == examId);
            if (exam == null)
            {
                return NotFound();
            }
            string examName = exam.ExamName;
            string level = exam.TestLevel;

            var employeeAnswers = await _context.ExamTrialRunAnswers.Where(a => a.ExamTrialRunId == examId).ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("data");

                worksheet.Cell(1, 1).Value = examName + " - " + level;
                worksheet.Range("A1:I1").Merge().Style.Font.Bold = true;
                worksheet.Range("A1:I1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                worksheet.Cell(2, 1).Value = "STT";
                worksheet.Cell(2, 2).Value = "MNV";
                worksheet.Cell(2, 3).Value = "Họ và tên";
                worksheet.Cell(2, 4).Value = "Câu trả lời";
                worksheet.Cell(2, 5).Value = "Đúng";
                worksheet.Cell(2, 6).Value = "Sai";
                worksheet.Cell(2, 7).Value = "Liệt";
                worksheet.Cell(2, 8).Value = "Ghi chú";
                worksheet.Cell(2, 9).Value = "Ngày làm bài";

                int row = 3;
                int stt = 1;
                foreach (var answer in employeeAnswers)
                {
                    worksheet.Cell(row, 1).Value = stt++;
                    worksheet.Cell(row, 2).Value = answer.EmployeeCode;
                    worksheet.Cell(row, 3).Value = answer.EmployeeName;
                    worksheet.Cell(row, 4).Value = answer.ListAnswer;
                    worksheet.Cell(row, 5).Value = answer.MultipleChoiceCorrect + answer.EssayCorrect;
                    worksheet.Cell(row, 6).Value = answer.MultipleChoiceInCorrect + answer.EssayInCorrect;
                    worksheet.Cell(row, 7).Value = answer.MultipleChoiceFail + answer.EssayFail;
                    worksheet.Cell(row, 8).Value = answer.Note;
                    worksheet.Cell(row, 9).Value = answer.CreatedDate;
                    row++;
                }

                worksheet.Columns().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "KQKT_" + examName.Replace(" ", "_") + ".xlsx"
                    );
                }
            }
        }

        [Authorize(Policy = "EditResultExam")]
        public async Task<IActionResult> Show(long examId, int isShow)
        {
            if (examId == null || isShow == null)
            {
                return NotFound();
            }
            var recordsToUpdate = await _context.ExamTrialRunAnswers.Where(e => e.ExamTrialRunId == examId).ToListAsync();
            if (recordsToUpdate == null)
            {
                return NotFound();
            }

            foreach (var record in recordsToUpdate)
            {
                record.IsShow = isShow;
                record.UpdatedDate = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            }

            try
            {
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Cập nhật thành công!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Detail), new { examid = examId });
        }
    }
}
