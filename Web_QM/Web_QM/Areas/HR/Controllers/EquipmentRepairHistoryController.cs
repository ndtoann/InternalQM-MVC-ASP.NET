using DocumentFormat.OpenXml.Bibliography;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Hosting;
using System.Reflection.PortableExecutable;
using Web_QM.Models;

namespace Web_QM.Areas.HR.Controllers
{
    [Area("HR")]
    public class EquipmentRepairHistoryController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public EquipmentRepairHistoryController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewMachineRepair")]
        public async Task<IActionResult> Index(string key, int isComplete = -1)
        {
            var res = await _context.EquipmentRepairHistories.AsNoTracking()
                .Where(m =>
                    (isComplete == -1 ||
                     (isComplete == 0 && m.CompletionDate == null) ||
                     (isComplete == 1 && m.CompletionDate != null)) &&
                    (string.IsNullOrEmpty(key) || m.EquipmentCode.ToLower().Contains(key.ToLower()) ||
                     m.EquipmentName.ToLower().Contains(key.ToLower()) ||
                     m.ErrorCondition.ToLower().Contains(key.ToLower()))
                )
                .Take(1000)
                .OrderByDescending(o => o.DateMonth)
                .ToListAsync();

            var statusOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "-1", Text = "Tất cả" },
                new SelectListItem { Value = "1", Text = "Đã sửa" },
                new SelectListItem { Value = "0", Text = "Chưa sửa" }
            };

            ViewData["IsCompleteList"] = new SelectList(statusOptions, "Value", "Text", isComplete);
            ViewBag.KeySearch = key;
            return View(res);
        }

        [Authorize(Policy = "AddMachineRepair")]
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

        [Authorize(Policy = "AddMachineRepair")]
        [HttpPost]
        public async Task<IActionResult> Add(EquipmentRepairHistory equipmentRepairHistory)
        {
            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewData["Departments"] = departmentsList;
            if (!ModelState.IsValid)
            {
                return View(equipmentRepairHistory);
            }
            try
            {
                _context.EquipmentRepairHistories.Add(equipmentRepairHistory);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!";
                return View(equipmentRepairHistory);
            }
        }

        [Authorize(Policy = "EditMachineRepair")]
        public async Task<IActionResult> Edit(long id)
        {
            if(id == null)
            {
                return NotFound();
            }
            var mrhToEdit = await _context.EquipmentRepairHistories.FindAsync(id);
            if (mrhToEdit == null)
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
            return View(mrhToEdit);
        }

        [Authorize(Policy = "EditMachineRepair")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, EquipmentRepairHistory equipmentRepairHistory)
        {
            if (id != equipmentRepairHistory.Id)
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
            if (!ModelState.IsValid)
            {
                return View(equipmentRepairHistory);
            }
            var mrhToEdit = await _context.EquipmentRepairHistories.AsNoTracking().FirstOrDefaultAsync(m => m.Id == equipmentRepairHistory.Id);
            if(mrhToEdit == null)
            {
                return NotFound();
            }
            try
            {
                _context.EquipmentRepairHistories.Update(equipmentRepairHistory);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!";
                return View(equipmentRepairHistory);
            }
        }

        [Authorize(Policy = "DeleteMachineRepair")]
        public async Task<IActionResult> Delete(long id)
        {
            if(id == null)
            {
                return NotFound();
            }
            var mrhToDelete = await _context.EquipmentRepairHistories.FindAsync(id);
            if(mrhToDelete == null)
            {
                return NotFound();
            }
            try
            {
                await _context.ReplacementEquipmentAndSupplies
                    .Where(a => a.EquipmentRepairId == id)
                    .ExecuteDeleteAsync();

                _context.EquipmentRepairHistories.Remove(mrhToDelete);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa lịch sử sửa chữa thiết bị!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }
            return RedirectToAction(nameof(Index));
        }

        [Authorize(Policy = "AddMachineRepair")]
        public async Task<IActionResult> GetEquipmentAndSupplies(long equipmentRepairId)
        {
            var list = await _context.ReplacementEquipmentAndSupplies
                               .Where(e => e.EquipmentRepairId == equipmentRepairId)
                               .ToListAsync();

            return Json(list);
        }

        [Authorize(Policy = "AddMachineRepair")]
        [HttpPost]
        public async Task<IActionResult> AddEquipmentAndSupplies([FromForm] ReplacementEquipmentAndSupplies model, IFormFile? FilePdf)
        {
            try
            {
                if (FilePdf != null && FilePdf.Length > 0)
                {
                    const int MaxFileSize = 5 * 1024 * 1024;
                    if (FilePdf.Length > MaxFileSize)
                    {
                        return Json(new { success = false, message = "Kích thước file không được vượt quá 5MB." });
                    }

                    string fileExtension = Path.GetExtension(FilePdf.FileName).ToLowerInvariant();
                    if (string.IsNullOrEmpty(fileExtension) || fileExtension != ".pdf")
                    {
                        return Json(new { success = false, message = "File tải lên phải là định dạng PDF (.pdf)." });
                    }

                    string uploadFolder = Path.Combine(_env.WebRootPath, "files", "equipments");
                    if (!Directory.Exists(uploadFolder))
                    {
                        Directory.CreateDirectory(uploadFolder);
                    }
                    string uniqueFileName = Guid.NewGuid().ToString() + "_" + FilePdf.FileName;
                    string filePath = Path.Combine(uploadFolder, uniqueFileName);

                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        await FilePdf.CopyToAsync(fileStream);
                    }
                    model.FilePdf = uniqueFileName;
                }

                _context.ReplacementEquipmentAndSupplies.Add(model);
                await _context.SaveChangesAsync();

                return Json(new { success = true, message = "Thêm mới thành công!" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Lỗi hệ thống khi lưu file hoặc dữ liệu!" });
            }
        }

        [Authorize(Policy = "AddMachineRepair")]
        public async Task<IActionResult> DeleteEquipmentAndSupplies(long id)
        {
            var item = await _context.ReplacementEquipmentAndSupplies.FindAsync(id);
            if (item == null)
            {
                return Json(new { success = false, message = "Không tìm thấy vật tư cần xóa!" });
            }

            try
            {
                if (!string.IsNullOrEmpty(item.FilePdf))
                {
                    string uploadFolder = Path.Combine(_env.WebRootPath, "files", "equipments");
                    string fullPath = Path.Combine(uploadFolder, item.FilePdf);

                    if (System.IO.File.Exists(fullPath))
                    {
                        System.IO.File.Delete(fullPath);
                    }
                }
                _context.ReplacementEquipmentAndSupplies.Remove(item);
                await _context.SaveChangesAsync();

                return Json(new { success = true, message = "Xóa thành công!" });
            }
            catch (Exception)
            {
                return Json(new { success = false, message = "Lỗi khi xóa vật tư hoặc file!" });
            }
        }

        [Authorize(Policy = "ViewMachineRepair")]
        public async Task<IActionResult> EquipmentReplace()
        {
            return View();
        }

        [Authorize(Policy = "ViewMachineRepair")]
        public async Task<IActionResult> GetReplacementDetails(string searchItem)
        {
            var result = await _context.ReplacementEquipmentAndSupplies
                .Where(r =>
                    string.IsNullOrEmpty(searchItem) ||
                    r.EquipmentAndlSupplies.ToLower().Contains(searchItem.Trim().ToLower())
                )
                .Join(
                    _context.EquipmentRepairHistories,
                    replacement => replacement.EquipmentRepairId,
                    repair => repair.Id,
                    (replacement, repair) => new
                    {
                        replacement.Id,
                        replacement.EquipmentAndlSupplies,
                        replacement.FilePdf,
                        replacement.EquipmentRepairId,
                        repair.EquipmentName,
                        repair.EquipmentCode,
                        repair.DateMonth
                    }
                )
                .ToListAsync();

            return Json(result);
        }
    }
}
