using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class PermissionController : Controller
    {
        private readonly QMContext _context;

        public PermissionController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ViewPermission")]
        public async Task<IActionResult> Index()
        {
            var res = await _context.Permissions.AsNoTracking().ToListAsync();

            return View(res);
        }

        [Authorize(Policy = "AddPermission")]
        public async Task<IActionResult> Add()
        {
            return View();
        }

        [Authorize(Policy = "AddPermission")]
        [HttpPost]
        public async Task<IActionResult> Add(Permission permission)
        {
            if (!ModelState.IsValid)
            {
                return View(permission);
            }
            bool isDuplicate = await IsDuplicateRole(permission.ClaimValue, permission.Description, 0);
            if (isDuplicate)
            {
                ModelState.AddModelError("ClaimValue", "Phân quyền đã tồn tại");
                return View(permission);
            }
            try
            {
                _context.Permissions.Add(permission);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(permission);
            }
        }

        [Authorize(Policy = "EditPermission")]
        public async Task<IActionResult> Edit(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var permissionToEdit = await _context.Permissions.FirstOrDefaultAsync(x => x.Id == id);
            if (permissionToEdit == null)
            {
                return NotFound();
            }
            return View(permissionToEdit);
        }

        [Authorize(Policy = "EditPermission")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, Permission permission)
        {
            if (id != permission.Id)
            {
                return NotFound();
            }
            if (!ModelState.IsValid)
            {
                return View(permission);
            }
            bool isDuplicate = await IsDuplicateRole(permission.ClaimValue, permission.Description, permission.Id);
            if (isDuplicate)
            {
                ModelState.AddModelError("ClaimValue", "Phân quyền đã tồn tại");
                return View(permission);
            }
            try
            {
                _context.Permissions.Update(permission);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(permission);
            }
        }

        [Authorize(Policy = "DeletePermission")]
        public async Task<IActionResult> Delete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var permissionToDelete = await _context.Permissions.FirstOrDefaultAsync(x => x.Id == id);

            if (permissionToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy phân quyền!";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                await _context.AccountPermissions
                     .Where(a => a.PermissionId == id)
                     .ExecuteDeleteAsync();

                _context.Permissions.Remove(permissionToDelete);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa thành công!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }
            return RedirectToAction(nameof(Index));
        }

        private async Task<bool> IsDuplicateRole(string claimValue, string description, long currentId = 0)
        {
            bool exists = await _context.Permissions
                .AnyAsync(p =>
                (
                    p.ClaimValue.ToLower() == claimValue.ToLower() ||
                    p.Description.ToLower() == description.ToLower()
                )
                && p.Id != currentId
            );
            return exists;
        }
    }
}
