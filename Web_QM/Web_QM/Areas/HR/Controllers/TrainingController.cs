using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;

namespace Web_QM.Areas.HR.Controllers
{
    [Area("HR")]
    [Authorize(Policy = "DataEmpl")]
    public class TrainingController : Controller
    {
        private readonly QMContext _context;

        public TrainingController( QMContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> Index()
        {
            var allTrainings = await _context.Trainings.AsNoTracking().ToListAsync();

            var generalTrainings = allTrainings.Where(t => t.Type == "General").ToList();
            var cncTrainings = allTrainings.Where(t => t.Type == "CNC").ToList();

            ViewBag.GeneralTrainings = generalTrainings;
            ViewBag.CNCTrainings = cncTrainings;
            return View();
        }

        public async Task<IActionResult> Add()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Add(Training training)
        {
            if (!ModelState.IsValid)
            {
                return View(training);
            }
            try
            {
                training.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Trainings.Add(training);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";

                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(training);
            }
        }

        public async Task<IActionResult> Edit(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var trainingToEdit = await _context.Trainings.FirstOrDefaultAsync(x => x.Id == id);
            if (trainingToEdit == null)
            {
                return NotFound();
            }

            return View(trainingToEdit);
        }

        [HttpPost]
        public async Task<IActionResult> Edit(long id, Training training)
        {
            if (id != training.Id)
            {
                return NotFound();
            }
            if (!ModelState.IsValid)
            {
                return View(training);
            }
            try
            {
                training.UpdatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Trainings.Update(training);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";

                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(training);
            }
        }

        public async Task<IActionResult> Delete(long id)
        {
            if(id == null)
            {
                return NotFound();
            }

            var trainingToDelete = await _context.Trainings.FirstOrDefaultAsync(x => x.Id == id);

            if (trainingToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy hạng mục này!";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                _context.Trainings.Remove(trainingToDelete);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Đã xóa thành công!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Index));
        }
    }
}
