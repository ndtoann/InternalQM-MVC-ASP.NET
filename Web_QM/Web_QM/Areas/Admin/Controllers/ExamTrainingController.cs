using ExcelDataReader;
using System.Globalization;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;
using Web_QM.Models.ViewModels;
using System.Data;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Office.SpreadSheetML.Y2023.MsForms;

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class ExamTrainingController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public ExamTrainingController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewExam")]
        public async Task<IActionResult> Index(string name, int isActive = -1)
        {
            var res = await _context.ExamTrainings
                                .AsNoTracking()
                                .Where(m => (isActive == -1 || m.IsActive == isActive)
                                && (string.IsNullOrEmpty(name) || m.ExamName.ToLower().Contains(name.ToLower())))
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

            var res = await _context.ExamTrainings.AsNoTracking().FirstOrDefaultAsync(m => m.Id == id);
            if (res == null)
            {
                return NotFound();
            }

            var listQuestion = await _context.QuestionTrainings.AsNoTracking().Where(q => q.ExamTrainingId == id).OrderBy(o => o.DisplayOrder).ToListAsync();
            var viewData = new ExamTrainingDetailView
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
        public async Task<IActionResult> Add(ExamTraining exam)
        {
            if (!ModelState.IsValid)
            {
                return View(exam);
            }
            var existExam = ExamExists(exam.ExamName);
            if (existExam)
            {
                ModelState.AddModelError("ExamName", "Tên bài kiểm tra đã tồn tại, vui lòng nhập tên khác!");
                return View(exam);
            }
            try
            {
                exam.EssayQuestion = ProcessBase64Images(exam.EssayQuestion);
                exam.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.ExamTrainings.Add(exam);
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

            var examToEdit = await _context.ExamTrainings.AsNoTracking().FirstOrDefaultAsync(m => m.Id == id);
            if (examToEdit == null)
            {
                return NotFound();
            }
            return View(examToEdit);
        }

        [Authorize(Policy = "EditExam")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, ExamTraining exam)
        {
            if (id != exam.Id)
            {
                return NotFound();
            }
            if (!ModelState.IsValid)
            {
                return View(exam);
            }
            var oldExam = await _context.ExamTrainings.AsNoTracking().FirstOrDefaultAsync(q => q.Id == exam.Id);
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
            try
            {
                if (!string.IsNullOrEmpty(oldExam.EssayQuestion) && oldExam.EssayQuestion.Contains("<img"))
                {
                    DeleteOldImages(oldExam.EssayQuestion, exam.EssayQuestion);
                }
                exam.EssayQuestion = ProcessBase64Images(exam.EssayQuestion);
                exam.UpdatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.ExamTrainings.Update(exam);
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
            return _context.ExamTrainings.Any(e => (e.Id != id) && (e.ExamName == examName));
        }

        [Authorize(Policy = "DeleteExam")]
        public async Task<IActionResult> Delete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var examToDelete = await _context.ExamTrainings.FindAsync(id);

            if (examToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy bài kiểm tra này!";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                var imageUrls = new List<string>();
                var regex = new Regex(@"src=""/imgs/uploads/questions/(?<fileName>.+?)""");

                if (!string.IsNullOrEmpty(examToDelete.EssayQuestion))
                {
                    var matches = regex.Matches(examToDelete.EssayQuestion);
                    foreach (Match match in matches)
                    {
                        imageUrls.Add($"/imgs/uploads/questions/{match.Groups["fileName"].Value}");
                    }
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

                _context.ExamTrainings.Remove(examToDelete);
                await _context.SaveChangesAsync();

                await _context.QuestionTrainings
                     .Where(q => q.ExamTrainingId == id)
                     .ExecuteDeleteAsync();

                await _context.ExamTrainingAnswers
                     .Where(e => e.ExamTrainingId == id)
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
        public async Task<IActionResult> AddQuestion(QuestionTraining question)
        {
            if (await ExistAnswer(question.ExamTrainingId))
            {
                TempData["ErrorMessage"] = "Bài kiểm tra đã được làm, vui lòng không sửa!";
                return RedirectToAction(nameof(Detail), new { id = question.ExamTrainingId });
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
                _context.QuestionTrainings.Add(question);
                await _context.SaveChangesAsync();
                question.DisplayOrder = (int)question.Id;
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Detail), new { id = question.ExamTrainingId });
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
            if (id == null)
            {
                return NotFound();
            }

            var questionToEdit = await _context.QuestionTrainings.FirstOrDefaultAsync(m => m.Id == id);
            if (questionToEdit == null)
            {
                return NotFound();
            }
            if (await ExistAnswer(questionToEdit.ExamTrainingId))
            {
                TempData["ErrorMessage"] = "Bài kiểm tra đã được làm, vui lòng không sửa!";
                return RedirectToAction(nameof(Detail), new { id = questionToEdit.ExamTrainingId });
            }
            ViewBag.ExamId = questionToEdit.ExamTrainingId;
            return View(questionToEdit);
        }

        [Authorize(Policy = "UpdateQuestion")]
        [HttpPost]
        public async Task<IActionResult> EditQuestion(QuestionTraining question)
        {
            if (await ExistAnswer(question.ExamTrainingId))
            {
                TempData["ErrorMessage"] = "Bài kiểm tra đã được làm, vui lòng không sửa!";
                return RedirectToAction(nameof(Detail), new { id = question.ExamTrainingId });
            }
            if (!ModelState.IsValid)
            {
                return View(question);
            }
            var oldQuestion = await _context.QuestionTrainings.AsNoTracking().FirstOrDefaultAsync(q => q.Id == question.Id);
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
                _context.QuestionTrainings.Update(question);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Detail), new { id = question.ExamTrainingId });
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
            // Regex để tìm các đường dẫn ảnh đã lưu
            var regex = new Regex(@"src=""/imgs/uploads/questions/(?<fileName>.+?)""");

            var oldMatches = regex.Matches(oldContent);
            var newMatches = regex.Matches(newContent);

            // Tạo danh sách các đường dẫn ảnh mới để so sánh
            var newImageUrls = new HashSet<string>();
            foreach (Match match in newMatches)
            {
                newImageUrls.Add(match.Groups["fileName"].Value);
            }

            // So sánh và xóa ảnh cũ không còn được sử dụng
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
            var questionToDelete = await _context.QuestionTrainings.FindAsync(id);

            if (questionToDelete == null)
            {
                return NotFound();
            }
            if (await ExistAnswer(questionToDelete.ExamTrainingId))
            {
                TempData["ErrorMessage"] = "Bài kiểm tra đã được làm, vui lòng không sửa!";
                return RedirectToAction(nameof(Detail), new { id = questionToDelete.ExamTrainingId });
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

                _context.QuestionTrainings.Remove(questionToDelete);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa câu hỏi!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Detail), new { id = questionToDelete.ExamTrainingId });
        }

        private List<string> ExtractImageUrlsFromQuestion(QuestionTraining question)
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
            var originalExam = await _context.ExamTrainings.FindAsync(examId);

            if (originalExam == null)
            {
                return NotFound();
            }

            var originalQuestions = await _context.QuestionTrainings
                                                  .Where(q => q.ExamTrainingId == examId)
                                                  .AsNoTracking()
                                                  .ToListAsync();

            var newExam = new ExamTraining
            {
                ExamName = originalExam.ExamName + " - Copy",
                DurationMinute = originalExam.DurationMinute,
                TlTotal = originalExam.TlTotal,
                IsActive = originalExam.IsActive,
                EssayQuestion = DuplicateImages(originalExam.EssayQuestion),
                CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy")
            };

            _context.ExamTrainings.Add(newExam);
            await _context.SaveChangesAsync();

            var newQuestions = new List<QuestionTraining>();
            foreach (var question in originalQuestions)
            {
                var newQuestion = new QuestionTraining
                {
                    QuestionText = DuplicateImages(question.QuestionText),
                    OptionA = DuplicateImages(question.OptionA),
                    OptionB = DuplicateImages(question.OptionB),
                    OptionC = DuplicateImages(question.OptionC),
                    OptionD = DuplicateImages(question.OptionD),
                    ExamTrainingId = newExam.Id,
                    CorrectOption = question.CorrectOption,
                    CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy")
                };
                newQuestions.Add(newQuestion);
            }

            _context.QuestionTrainings.AddRange(newQuestions);
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

        [Authorize(Policy = "UpdateQuestion")]
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

            var questionList = new List<QuestionTraining>();

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
                                var question = new QuestionTraining
                                {
                                    QuestionText = row["QuestionText"].ToString(),
                                    OptionA = row["OptionA"].ToString(),
                                    OptionB = row["OptionB"].ToString(),
                                    OptionC = row["OptionC"].ToString(),
                                    OptionD = row["OptionD"].ToString(),
                                    CorrectOption = row["CorrectOption"].ToString(),
                                    ExamTrainingId = examId,
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
                await _context.QuestionTrainings.AddRangeAsync(questionList);
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

            var question = await _context.QuestionTrainings.FirstOrDefaultAsync(q => q.Id == id);
            if (question == null)
            {
                return Json(new { success = false, message = "Không tìm thấy câu hỏi." });
            }
            if (await ExistAnswer(question.ExamTrainingId))
            {
                return Json(new { success = false, message = "Bài kiểm tra đã được làm, vui lòng không sửa!" });
            }
            try
            {
                question.CorrectOption = correctOption;
                _context.QuestionTrainings.Update(question);
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
            var currentQuestion = await _context.QuestionTrainings.FirstOrDefaultAsync(q => q.Id == questionId);
            if (currentQuestion == null)
            {
                return Json(new { success = false, message = "Không tìm thấy câu hỏi." });
            }
            if (await ExistAnswer(currentQuestion.ExamTrainingId))
            {
                return Json(new { success = false, message = "Bài kiểm tra đã được làm, vui lòng không sửa!" });
            }

            var questionsInExam = await _context.QuestionTrainings
                                               .Where(q => q.ExamTrainingId == currentQuestion.ExamTrainingId)
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

            _context.QuestionTrainings.Update(currentQuestion);
            _context.QuestionTrainings.Update(targetQuestion);

            await _context.SaveChangesAsync();

            return Json(new { success = true, message = "Thứ tự câu hỏi đã được cập nhật thành công." });
        }

        [Authorize(Policy = "UpdateQuestion")]
        [HttpPost]
        public async Task<IActionResult> UpdateQuestionPosition(long questionId, long targetQuestionId)
        {
            var currentQuestion = await _context.QuestionTrainings.FirstOrDefaultAsync(q => q.Id == questionId);
            if (await ExistAnswer(currentQuestion.ExamTrainingId))
            {
                return Json(new { success = false, message = "Bài kiểm tra đã được làm, vui lòng không sửa!" });
            }
            var targetQuestion = await _context.QuestionTrainings.FirstOrDefaultAsync(q => q.Id == targetQuestionId);

            if (currentQuestion == null || targetQuestion == null)
            {
                return Json(new { success = false, message = "Không tìm thấy một trong hai câu hỏi." });
            }

            if (currentQuestion.ExamTrainingId != targetQuestion.ExamTrainingId)
            {
                return Json(new { success = false, message = "Không thể hoán đổi câu hỏi từ các bài thi khác nhau." });
            }

            var tempOrder = currentQuestion.DisplayOrder;
            currentQuestion.DisplayOrder = targetQuestion.DisplayOrder;
            targetQuestion.DisplayOrder = tempOrder;

            _context.QuestionTrainings.Update(currentQuestion);
            _context.QuestionTrainings.Update(targetQuestion);

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
            var hasAnswer = await _context.ExamTrainingAnswers.AnyAsync(a => a.ExamTrainingId == examId);
            return hasAnswer;
        }
    }
}
