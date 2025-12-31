using ExcelDataReader;
using System.Globalization;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;
using Web_QM.Models.ViewModels;
using System.Data;
using Microsoft.AspNetCore.Hosting;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.Threading.Tasks;

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class ExamController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public ExamController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewExam")]
        public async Task<IActionResult> Index(string name, string department, int isActive = -1)
        {
            var res = await _context.ExamPeriodics
                                .AsNoTracking()
                                .Where(m => (isActive == -1 || m.IsActive == isActive)
                                && (string.IsNullOrEmpty(name) || m.ExamName.ToLower().Contains(name.ToLower()))
                                && (string.IsNullOrEmpty(department) || m.ExamName.ToLower().Contains(department.ToLower())))
                                .OrderByDescending(e => e.Id)
                                .Take(200)
                                .ToListAsync();

            var statusOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "-1", Text = "Tất cả" },
                new SelectListItem { Value = "1", Text = "Sẵn sàng" },
                new SelectListItem { Value = "0", Text = "Tạm dừng" }
            };
            ViewData["IsActiveList"] = new SelectList(statusOptions, "Value", "Text", isActive);
            var departments = await _context.Departments.AsNoTracking().ToListAsync();
            var departmentsList = new List<SelectListItem>
            {
                new SelectListItem { Value = "", Text = "Tất cả" },
                new SelectListItem { Value = "Kỹ thuật", Text = "Kỹ thuật" },
                new SelectListItem { Value = "Kế hoạch", Text = "Kế hoạch" },
                new SelectListItem { Value = "HC-KT", Text = "HC-KT" },
                new SelectListItem { Value = "KCS", Text = "KCS" },
                new SelectListItem { Value = "Kho", Text = "Kho" },
                new SelectListItem { Value = "CNC", Text = "CNC" },
                new SelectListItem { Value = "Taro", Text = "Taro" },
                new SelectListItem { Value = "Bavia", Text = "Bavia" },
                new SelectListItem { Value = "Hàn-Cưa", Text = "Hàn-Cưa" },
                new SelectListItem { Value = "Phay", Text = "Phay" },
                new SelectListItem { Value = "Tiện", Text = "Tiện" },
                new SelectListItem { Value = "Đánh bóng", Text = "Đánh bóng" },
            };
            ViewData["Departments"] = new SelectList(departmentsList, "Value", "Text", department);
            ViewBag.Name = name;

            return View(res);
        }

        [Authorize(Policy = "ViewExam")]
        public async Task<IActionResult> Detail(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var res = await _context.ExamPeriodics.AsNoTracking().FirstOrDefaultAsync(m => m.Id == id);
            if (res == null)
            {
                return NotFound();
            }

            var listQuestion = await _context.Questions.AsNoTracking().Where(q => q.ExamId == id).OrderBy(o => o.DisplayOrder).ToListAsync();
            var viewData = new ExamDetailView
            {
                exam = res,
                questions = listQuestion
            };
            return View(viewData);
        }

        [Authorize(Policy = "AddExam")]
        public async Task<IActionResult> Add()
        {
            return View();
        }

        [Authorize(Policy = "AddExam")]
        [HttpPost]
        public async Task<IActionResult> Add(ExamPeriodic exam, IFormFile essayPdf)
        {
            var existExam = ExamExists(exam.ExamName);
            if (existExam)
            {
                ModelState.AddModelError("ExamName", "Tên bài kiểm tra đã tồn tại, vui lòng nhập tên khác!");
                return View(exam);
            }
            const long MaxFileSize = 20 * 1024 * 1024;
            if (essayPdf != null && essayPdf.Length > 0)
            {
                if (essayPdf.Length > MaxFileSize)
                {
                    ModelState.AddModelError("EssayQuestion", "Kích thước file PDF không được vượt quá 20MB!");
                    return View(exam);
                }
                exam.EssayQuestion = await SavePdfFile(essayPdf);
                if (string.IsNullOrEmpty(exam.EssayQuestion))
                {
                    ModelState.AddModelError("EssayQuestion", "Vui lòng chọn file PDF!");
                    return View(exam);
                }
            }
            try
            {
                exam.CreatedDate = DateOnly.FromDateTime(DateTime.Now);
                _context.ExamPeriodics.Add(exam);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(exam);
            }
        }

        [Authorize(Policy = "EditExam")]
        public async Task<IActionResult> Edit(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var examToEdit = await _context.ExamPeriodics.AsNoTracking().FirstOrDefaultAsync(m => m.Id == id);
            if (examToEdit == null)
            {
                return NotFound();
            }
            return View(examToEdit);
        }

        [Authorize(Policy = "EditExam")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, ExamPeriodic exam, IFormFile essayPdf)
        {
            if (id != exam.Id)
            {
                return NotFound();
            }
            var oldExam = await _context.ExamPeriodics.AsNoTracking().FirstOrDefaultAsync(q => q.Id == exam.Id);
            if (oldExam == null)
            {
                return NotFound();
            }
            var existExam = ExamExists(exam.ExamName, exam.Id);
            if (existExam)
            {
                ModelState.AddModelError("ExamName", "Tên bài kiểm tra đã tồn tại, vui lòng nhập tên khác!");
                return View(exam);
            }
            const long MaxFileSize = 20 * 1024 * 1024;
            if (essayPdf != null && essayPdf.Length > 0)
            {
                if (essayPdf.Length > MaxFileSize)
                {
                    ModelState.AddModelError("EssayQuestion", "Kích thước file PDF không được vượt quá 20MB!");
                    return View(exam);
                }
                exam.EssayQuestion = await SavePdfFile(essayPdf);
                if (string.IsNullOrEmpty(exam.EssayQuestion))
                {
                    ModelState.AddModelError("EssayQuestion", "Vui lòng chọn file PDF!");
                    return View(exam);
                }
                DeletePdfFile(oldExam.EssayQuestion);
            }
            else
            {
                exam.EssayQuestion = oldExam.EssayQuestion;
            }
            try
            {
                exam.UpdatedDate = DateOnly.FromDateTime(DateTime.Now);
                _context.ExamPeriodics.Update(exam);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(exam);
            }
        }

        private bool ExamExists(string examName, long? id = 0)
        {
            return _context.ExamPeriodics.Any(e => (e.Id != id) && (e.ExamName == examName));
        }

        [Authorize(Policy = "DeleteExam")]
        public async Task<IActionResult> Delete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var examToDelete = await _context.ExamPeriodics.FindAsync(id);

            if (examToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy bài kiểm tra này!";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                var imageUrls = new List<string>();

                if (!string.IsNullOrEmpty(examToDelete.EssayQuestion))
                {
                    DeletePdfFile(examToDelete.EssayQuestion);
                }

                foreach (var imageUrl in imageUrls)
                {
                    var fileName = Path.GetFileName(imageUrl);
                    var filePath = Path.Combine(_env.WebRootPath, "imgs", "uploads", "questions", fileName);

                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }

                _context.ExamPeriodics.Remove(examToDelete);
                await _context.SaveChangesAsync();

                await _context.Questions
                     .Where(q => q.ExamId == id)
                     .ExecuteDeleteAsync();

                await _context.ExamPeriodicAnswers
                     .Where(e => e.ExamId == id)
                     .ExecuteDeleteAsync();

                TempData["SuccessMessage"] = "Đã xóa bài kiểm tra!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Index));
        }

        //xử lý câu hỏi
        [Authorize(Policy = "UpdateQuestion")]
        public async Task<IActionResult> AddQuestion(long examid)
        {
            if (await ExistAnswer(examid))
            {
                TempData["ErrorMessage"] = "Bài kiểm tra đã được làm, vui lòng không sửa!";
                return RedirectToAction(nameof(Detail), new { id = examid });
            }
            ViewBag.ExamId = examid;
            return View();
        }

        [Authorize(Policy = "UpdateQuestion")]
        [HttpPost]
        public async Task<IActionResult> AddQuestion(Question question)
        {
            if (await ExistAnswer(question.ExamId))
            {
                TempData["ErrorMessage"] = "Bài kiểm tra đã được làm, vui lòng không sửa!";
                return RedirectToAction(nameof(Detail), new { id = question.ExamId });
            }
            if (!ModelState.IsValid)
            {
                return View(question);
            }
            try
            {
                question.QuestionText = ProcessBase64Images(question.QuestionText);
                question.OptionA = ProcessBase64Images(question.OptionA);
                question.OptionB = ProcessBase64Images(question.OptionB);
                question.OptionC = ProcessBase64Images(question.OptionC);
                question.OptionD = ProcessBase64Images(question.OptionD);
                question.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Questions.Add(question);
                await _context.SaveChangesAsync();
                question.DisplayOrder = (int)question.Id;
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";

                return RedirectToAction(nameof(Detail), new {id = question.ExamId });
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(question);
            }
        }

        [Authorize(Policy = "UpdateQuestion")]
        public async Task<IActionResult> EditQuestion(long id)
        {
            if(id == null)
            {
                return NotFound();
            }

            var questionToEdit = await _context.Questions.FirstOrDefaultAsync(m => m.Id == id);
            if (questionToEdit == null)
            {
                return NotFound();
            }
            if (await ExistAnswer(questionToEdit.ExamId))
            {
                TempData["ErrorMessage"] = "Bài kiểm tra đã được làm, vui lòng không sửa!";
                return RedirectToAction(nameof(Detail), new { id = questionToEdit.ExamId });
            }
            ViewBag.ExamId = questionToEdit.ExamId;
            return View(questionToEdit);
        }

        [Authorize(Policy = "UpdateQuestion")]
        [HttpPost]
        public async Task<IActionResult> EditQuestion(Question question)
        {
            if (await ExistAnswer(question.ExamId))
            {
                TempData["ErrorMessage"] = "Bài kiểm tra đã được làm, vui lòng không sửa!";
                return RedirectToAction(nameof(Detail), new { id = question.ExamId });
            }
            if (!ModelState.IsValid)
            {
                return View(question);
            }
            var oldQuestion = await _context.Questions.AsNoTracking().FirstOrDefaultAsync(q => q.Id == question.Id);
            if (oldQuestion == null)
            {
                return NotFound();
            }

            try
            {
                question.QuestionText = ProcessBase64Images(question.QuestionText);
                question.OptionA = ProcessBase64Images(question.OptionA);
                question.OptionB = ProcessBase64Images(question.OptionB);
                question.OptionC = ProcessBase64Images(question.OptionC);
                question.OptionD = ProcessBase64Images(question.OptionD);

                DeleteOldImages(oldQuestion.QuestionText, question.QuestionText);
                DeleteOldImages(oldQuestion.OptionA, question.OptionA);
                DeleteOldImages(oldQuestion.OptionB, question.OptionB);
                DeleteOldImages(oldQuestion.OptionC, question.OptionC);
                DeleteOldImages(oldQuestion.OptionD, question.OptionD);

                question.DisplayOrder = oldQuestion.DisplayOrder;
                question.UpdatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Questions.Update(question);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";

                return RedirectToAction(nameof(Detail), new { id = question.ExamId });
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(question);
            }
        }

        private string ProcessBase64Images(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return content;
            }

            var regex = new Regex(@"src=""data:image/(?<type>.+?);base64,(?<data>[A-Za-z0-9+/=]+)""");
            var matches = regex.Matches(content);

            foreach (Match match in matches)
            {
                var base64Data = match.Groups["data"].Value;
                var imageType = match.Groups["type"].Value;

                try
                {
                    byte[] imageBytes = Convert.FromBase64String(base64Data);

                    var fileName = $"{Guid.NewGuid().ToString()}.{imageType}";
                    var filePath = Path.Combine("imgs", "uploads", "questions", fileName);
                    var fullPath = Path.Combine(_env.WebRootPath, filePath);

                    var directoryPath = Path.GetDirectoryName(fullPath);
                    if (!Directory.Exists(directoryPath))
                    {
                        Directory.CreateDirectory(directoryPath);
                    }

                    System.IO.File.WriteAllBytes(fullPath, imageBytes);

                    var newImageUrl = $"/{filePath.Replace("\\", "/")}";

                    content = content.Replace(match.Value, $"src=\"{newImageUrl}\"");
                }
                catch (Exception ex)
                {
                    continue;
                }
            }
            return content;
        }

        private void DeleteOldImages(string oldContent, string newContent)
        {
            var regex = new Regex(@"src=""/imgs/uploads/questions/(?<fileName>.+?)""");

            var oldMatches = regex.Matches(oldContent);
            var newMatches = regex.Matches(newContent);

            var newImageUrls = new HashSet<string>();
            foreach (Match match in newMatches)
            {
                newImageUrls.Add(match.Groups["fileName"].Value);
            }

            foreach (Match match in oldMatches)
            {
                var oldFileName = match.Groups["fileName"].Value;
                if (!newImageUrls.Contains(oldFileName))
                {
                    var oldFilePath = Path.Combine(_env.WebRootPath, "imgs", "uploads", "questions", oldFileName);
                    if (System.IO.File.Exists(oldFilePath))
                    {
                        System.IO.File.Delete(oldFilePath);
                    }
                }
            }
        }

        [Authorize(Policy = "UpdateQuestion")]
        public async Task<IActionResult> DeleteQuestion(long id)
        {
            var questionToDelete = await _context.Questions.FindAsync(id);

            if (questionToDelete == null)
            {
                return NotFound();
            }
            if (await ExistAnswer(questionToDelete.ExamId))
            {
                TempData["ErrorMessage"] = "Bài kiểm tra đã được làm, vui lòng không sửa!";
                return RedirectToAction(nameof(Detail), new { id = questionToDelete.ExamId });
            }

            try
            {
                var allImageUrls = ExtractImageUrlsFromQuestion(questionToDelete);

                foreach (var imageUrl in allImageUrls)
                {
                    var fileName = Path.GetFileName(imageUrl);
                    var filePath = Path.Combine(_env.WebRootPath, "imgs", "uploads", "questions", fileName);

                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                }

                _context.Questions.Remove(questionToDelete);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa câu hỏi!";

            }
            catch (Exception)
            {
                TempData["ErrorMessage"] = "Đã xảy ra lỗi khi xóa câu hỏi. Vui lòng thử lại!";
            }
            return RedirectToAction(nameof(Detail), new { id = questionToDelete.ExamId });
        }

        private List<string> ExtractImageUrlsFromQuestion(Question question)
        {
            var imageUrls = new List<string>();
            var regex = new Regex(@"src=""/imgs/uploads/questions/(?<fileName>.+?)""");

            if (!string.IsNullOrEmpty(question.QuestionText))
            {
                var matches = regex.Matches(question.QuestionText);
                foreach (Match match in matches)
                {
                    imageUrls.Add($"/imgs/uploads/questions/{match.Groups["fileName"].Value}");
                }
            }

            if (!string.IsNullOrEmpty(question.OptionA))
            {
                var matches = regex.Matches(question.OptionA);
                foreach (Match match in matches)
                {
                    imageUrls.Add($"/imgs/uploads/questions/{match.Groups["fileName"].Value}");
                }
            }

            if (!string.IsNullOrEmpty(question.OptionB))
            {
                var matches = regex.Matches(question.OptionB);
                foreach (Match match in matches)
                {
                    imageUrls.Add($"/imgs/uploads/questions/{match.Groups["fileName"].Value}");
                }
            }
            if (!string.IsNullOrEmpty(question.OptionC))
            {
                var matches = regex.Matches(question.OptionC);
                foreach (Match match in matches)
                {
                    imageUrls.Add($"/imgs/uploads/questions/{match.Groups["fileName"].Value}");
                }
            }
            if (!string.IsNullOrEmpty(question.OptionD))
            {
                var matches = regex.Matches(question.OptionD);
                foreach (Match match in matches)
                {
                    imageUrls.Add($"/imgs/uploads/questions/{match.Groups["fileName"].Value}");
                }
            }

            return imageUrls;
        }

        [Authorize(Policy = "AddExam")]
        public async Task<IActionResult> DuplicateExam(long examId)
        {
            var originalExam = await _context.ExamPeriodics.FindAsync(examId);

            if (originalExam == null)
            {
                return NotFound();
            }

            var originalQuestions = await _context.Questions
                                                 .Where(q => q.ExamId == examId)
                                                 .AsNoTracking()
                                                 .ToListAsync();
            string newEssayQuestionFileName = originalExam.EssayQuestion;
            if (!string.IsNullOrEmpty(originalExam.EssayQuestion))
            {
                string essayQuestionsDirectory = Path.Combine(_env.WebRootPath, "files", "essay_questions");
                string originalFilePath = Path.Combine(essayQuestionsDirectory, originalExam.EssayQuestion);
                if (System.IO.File.Exists(originalFilePath))
                {
                    string originalFileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalExam.EssayQuestion);
                    string extension = Path.GetExtension(originalExam.EssayQuestion);

                    newEssayQuestionFileName = $"{Guid.NewGuid().ToString("N")}{extension}";

                    string newFilePath = Path.Combine(essayQuestionsDirectory, newEssayQuestionFileName);

                    try
                    {
                        System.IO.File.Copy(originalFilePath, newFilePath);
                    }
                    catch (Exception ex)
                    {
                        TempData["ErrorMessage"] = "Lỗi khi copy file tự luận!";
                        return RedirectToAction(nameof(Index));
                    }
                }
                else
                {
                    newEssayQuestionFileName = null;
                }
            }
            var newExam = new ExamPeriodic
            {
                ExamName = originalExam.ExamName + " - Copy",
                DurationMinute = originalExam.DurationMinute,
                EssayQuestion = newEssayQuestionFileName,
                TlTotal = originalExam.TlTotal,
                IsActive = originalExam.IsActive,
                CreatedDate = DateOnly.FromDateTime(DateTime.Now)
            };

            _context.ExamPeriodics.Add(newExam);
            await _context.SaveChangesAsync();

            var newQuestions = new List<Question>();
            foreach (var question in originalQuestions)
            {
                var newQuestion = new Question
                {
                    QuestionText = DuplicateImages(question.QuestionText),
                    OptionA = DuplicateImages(question.OptionA),
                    OptionB = DuplicateImages(question.OptionB),
                    OptionC = DuplicateImages(question.OptionC),
                    OptionD = DuplicateImages(question.OptionD),
                    ExamId = newExam.Id,
                    CorrectOption = question.CorrectOption,
                    CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy")
                };
                newQuestions.Add(newQuestion);
            }

            _context.Questions.AddRange(newQuestions);
            await _context.SaveChangesAsync();

            TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
            return RedirectToAction(nameof(Index));
        }

        private string DuplicateImages(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return content;
            }

            var regex = new Regex(@"src=""/imgs/uploads/questions/(?<fileName>.+?)""");
            var matches = regex.Matches(content);

            foreach (Match match in matches)
            {
                var oldFileName = match.Groups["fileName"].Value;
                var newFileName = $"{Guid.NewGuid().ToString()}{Path.GetExtension(oldFileName)}";

                var oldFilePath = Path.Combine(_env.WebRootPath, "imgs", "uploads", "questions", oldFileName);
                var newFilePath = Path.Combine(_env.WebRootPath, "imgs", "uploads", "questions", newFileName);

                if (System.IO.File.Exists(oldFilePath))
                {
                    try
                    {
                        System.IO.File.Copy(oldFilePath, newFilePath);
                        var newImageUrl = $"/imgs/uploads/questions/{newFileName}";
                        content = content.Replace(match.Value, $"src=\"{newImageUrl}\"");
                    }
                    catch (Exception ex)
                    {
                        continue;
                    }
                }
            }
            return content;
        }

        [Authorize(Policy = "AddExam")]
        [HttpPost]
        public async Task<IActionResult> ImportExcelQuestion(IFormFile file, long examId)
        {
            if (await ExistAnswer(examId))
            {
                return BadRequest(new { Message = "Bài kiểm tra đã được làm, vui lòng không sửa!" });
            }
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { Message = "Vui lòng chọn một file Excel." });
            }

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var questionList = new List<Question>();

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

                        if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                        {
                            return BadRequest(new { Message = "File Excel không có dữ liệu." });
                        }

                        DataTable table = dataSet.Tables[0];
                        var requiredFields = new[] { "QuestionText", "OptionA", "OptionB", "OptionC", "OptionD", "CorrectOption" };

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            var row = table.Rows[i];
                            var missingFields = new List<string>();

                            foreach (var field in requiredFields)
                            {
                                if (!table.Columns.Contains(field) || row[field] == DBNull.Value || string.IsNullOrWhiteSpace(row[field].ToString()))
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
                                var question = new Question
                                {
                                    QuestionText = row["QuestionText"].ToString(),
                                    OptionA = row["OptionA"].ToString(),
                                    OptionB = row["OptionB"].ToString(),
                                    OptionC = row["OptionC"].ToString(),
                                    OptionD = row["OptionD"].ToString(),
                                    CorrectOption = row["CorrectOption"].ToString(),
                                    ExamId = examId,
                                    CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy")
                                };
                                questionList.Add(question);
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
                return StatusCode(500, new { Message = $"Có lỗi xảy ra khi đọc file: {ex.Message}" });
            }

            try
            {
                await _context.Questions.AddRangeAsync(questionList);
                await _context.SaveChangesAsync();

                foreach (var question in questionList)
                {
                    question.DisplayOrder = (int)Math.Min(question.Id, int.MaxValue);
                }

                await _context.SaveChangesAsync();

                return Ok(new { Message = $"Lưu dữ liệu thành công. Đã import {questionList.Count} câu hỏi." });
            }
            catch (DbUpdateException ex)
            {
                return StatusCode(500, new { Message = $"Có lỗi xảy ra khi lưu dữ liệu vào database: {ex.Message}" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { Message = $"Có lỗi xảy ra: {ex.Message}" });
            }
        }

        [Authorize(Policy = "UpdateQuestion")]
        [HttpPost]
        public async Task<IActionResult> UpdateCorrectOption(long id, string correctOption)
        {
            if (id <= 0 || string.IsNullOrEmpty(correctOption))
            {
                return Json(new { success = false, message = "Dữ liệu không hợp lệ." });
            }

            var question = await _context.Questions.FirstOrDefaultAsync(q => q.Id == id);
            if (question == null)
            {
                return Json(new { success = false, message = "Không tìm thấy câu hỏi." });
            }
            if (await ExistAnswer(question.ExamId))
            {
                return Json(new { success = false, message = "Bài kiểm tra đã được làm, vui lòng không sửa!" });
            }

            try
            {
                question.CorrectOption = correctOption;
                _context.Questions.Update(question);
                await _context.SaveChangesAsync();
                return Json(new { success = true, message = "Cập nhật đáp án thành công." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Đã xảy ra lỗi hệ thống khi cập nhật." });
            }
        }

        [Authorize(Policy = "UpdateQuestion")]
        [HttpPost]
        public async Task<IActionResult> UpdateQuestionOrder(long questionId, string direction)
        {
            var currentQuestion = await _context.Questions.FirstOrDefaultAsync(q => q.Id == questionId);
            if (currentQuestion == null)
            {
                return Json(new { success = false, message = "Không tìm thấy câu hỏi." });
            }
            if (await ExistAnswer(currentQuestion.ExamId))
            {
                return Json(new { success = false, message = "Bài kiểm tra đã được làm, vui lòng không sửa!" });
            }

            var questionsInExam = await _context.Questions
                                               .Where(q => q.ExamId == currentQuestion.ExamId)
                                               .OrderBy(q => q.DisplayOrder)
                                               .ToListAsync();

            var currentIndex = questionsInExam.FindIndex(q => q.Id == questionId);
            int newIndex;

            if (direction == "up")
            {
                if (currentIndex <= 0)
                {
                    return Json(new { success = false, message = "Câu hỏi đã ở vị trí đầu tiên." });
                }
                newIndex = currentIndex - 1;
            }
            else if (direction == "down")
            {
                if (currentIndex >= questionsInExam.Count - 1)
                {
                    return Json(new { success = false, message = "Câu hỏi đã ở vị trí cuối cùng." });
                }
                newIndex = currentIndex + 1;
            }
            else
            {
                return Json(new { success = false, message = "Hướng sắp xếp không hợp lệ." });
            }

            var targetQuestion = questionsInExam[newIndex];

            var tempOrder = currentQuestion.DisplayOrder;
            currentQuestion.DisplayOrder = targetQuestion.DisplayOrder;
            targetQuestion.DisplayOrder = tempOrder;

            _context.Questions.Update(currentQuestion);
            _context.Questions.Update(targetQuestion);

            await _context.SaveChangesAsync();

            return Json(new { success = true, message = "Thứ tự câu hỏi đã được cập nhật thành công." });
        }

        [Authorize(Policy = "UpdateQuestion")]
        [HttpPost]
        public async Task<IActionResult> UpdateQuestionPosition(long questionId, long targetQuestionId)
        {
            var currentQuestion = await _context.Questions.FirstOrDefaultAsync(q => q.Id == questionId);
            if (await ExistAnswer(currentQuestion.ExamId))
            {
                return Json(new { success = false, message = "Bài kiểm tra đã được làm, vui lòng không sửa!" });
            }
            var targetQuestion = await _context.Questions.FirstOrDefaultAsync(q => q.Id == targetQuestionId);

            if (currentQuestion == null || targetQuestion == null)
            {
                return Json(new { success = false, message = "Không tìm thấy một trong hai câu hỏi." });
            }

            if (currentQuestion.ExamId != targetQuestion.ExamId)
            {
                return Json(new { success = false, message = "Không thể hoán đổi câu hỏi từ các bài thi khác nhau." });
            }

            var tempOrder = currentQuestion.DisplayOrder;
            currentQuestion.DisplayOrder = targetQuestion.DisplayOrder;
            targetQuestion.DisplayOrder = tempOrder;

            _context.Questions.Update(currentQuestion);
            _context.Questions.Update(targetQuestion);

            await _context.SaveChangesAsync();
            return Json(new { success = true, message = "Đã cập nhật vị trí câu hỏi thành công." });
        }

        [Authorize(Policy = "UpdateQuestion")]
        public async Task<IActionResult> DownloadTemplate()
        {
            var filePath = Path.Combine(_env.WebRootPath, "templates", "import_Question1.xlsx");
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound("Không tìm thấy file mẫu.");
            }
            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = "import_Question1.xlsx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        private async Task<bool> ExistAnswer(long examId)
        {
            if (examId == null)
            {
                return false;
            }
            var hasAnswer = await _context.ExamPeriodicAnswers.AnyAsync(a => a.ExamId == examId);
            return hasAnswer;
        }

        private async Task<string> SavePdfFile(IFormFile file)
        {
            if (file.ContentType != "application/pdf")
            {
                return null;
            }

            var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files", "essay_questions");

            if (!Directory.Exists(uploadsFolder))
            {
                Directory.CreateDirectory(uploadsFolder);
            }

            var uniqueFileName = Guid.NewGuid().ToString("N");
            var fileName = $"{uniqueFileName}.pdf";

            var filePath = Path.Combine(uploadsFolder, fileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            return fileName;
        }

        private async Task<bool> DeletePdfFile(string fileName)
        {
            try
            {
                var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files", "essay_questions");
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
    }
}
