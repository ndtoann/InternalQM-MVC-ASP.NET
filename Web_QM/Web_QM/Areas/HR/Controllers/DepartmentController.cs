using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;

namespace Web_QM.Areas.HR.Controllers
{
    [Area("HR")]
    public class DepartmentController : Controller
    {
        private readonly QMContext _context;

        public DepartmentController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ViewDepartment")]
        public async Task<IActionResult> Index()
        {
            var res = await _context.Departments.AsNoTracking().ToListAsync();
            return View(res);
        }

        [Authorize(Policy = "AddDepartment")]
        public async Task<IActionResult> Add()
        {
            return View();
        }

        [Authorize(Policy = "AddDepartment")]
        [HttpPost]
        public async Task<IActionResult> Add(Department department)
        {
            if (!ModelState.IsValid)
            {
                return View(department);
            }
            try
            {
                _context.Departments.Add(department);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(department);
            }
        }

        [Authorize(Policy = "EditDepartment")]
        public async Task<IActionResult> Edit(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var departmentToEdit = await _context.Departments.FirstOrDefaultAsync(x => x.Id == id);
            if (departmentToEdit == null)
            {
                return NotFound();
            }

            return View(departmentToEdit);
        }

        [Authorize(Policy = "EditDepartment")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, Department department)
        {
            if (id != department.Id)
            {
                return NotFound();
            }
            if (!ModelState.IsValid)
            {
                return View(department);
            }
            try
            {
                _context.Departments.Update(department);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(department);
            }
        }

        [Authorize(Policy = "DeleteDepartment")]
        public async Task<IActionResult> Delete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var departmentToDelete = await _context.Departments.FirstOrDefaultAsync(x => x.Id == id);

            if (departmentToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy bộ phận!";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                _context.Departments.Remove(departmentToDelete);
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
