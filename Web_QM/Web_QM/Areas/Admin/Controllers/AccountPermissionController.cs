using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;
using Web_QM.Models.ViewModels;

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class AccountPermissionController : Controller
    {
        private readonly QMContext _context;

        public AccountPermissionController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "EmplPermission")]
        public async Task<IActionResult> Index()
        {
            var allPermissions = await _context.Permissions.OrderBy(p => p.Module).ThenBy(p => p.ClaimValue).ToListAsync();

            var model = new ManageAccountPermissionsViewModel
            {
                AvailableAccounts = await _context.Accounts.ToListAsync(),
                GroupedPermissions = allPermissions
                    .GroupBy(p => p.Module)
                    .ToDictionary(g => g.Key, g => g.ToList())
            };

            return View(model);
        }

        [Authorize(Policy = "EmplPermission")]
        public async Task<IActionResult> GetAccountPermissions(long accountId)
        {
            if (accountId <= 0)
            {
                return Json(new List<long>());
            }

            var assignedPermissionIds = await _context.AccountPermissions
                .Where(ap => ap.AccountId == accountId)
                .Select(ap => ap.PermissionId)
                .ToListAsync();

            return Json(assignedPermissionIds);
        }

        [Authorize(Policy = "EmplPermission")]
        [HttpPost]
        public async Task<IActionResult> UpdatePermissions([FromBody] PermissionUpdateModel model)
        {
            if (model.SelectedAccountId == 0)
            {
                return Json(new { success = false, message = "Vui lòng chọn một tài khoản!" });
            }

            try
            {
                var currentPermissions = await _context.AccountPermissions
                    .Where(ap => ap.AccountId == model.SelectedAccountId)
                    .ToListAsync();

                _context.AccountPermissions.RemoveRange(currentPermissions);

                if (model.AssignedPermissionIds != null && model.AssignedPermissionIds.Any())
                {
                    var newPermissions = model.AssignedPermissionIds.Select(permissionId => new AccountPermission
                    {
                        AccountId = model.SelectedAccountId,
                        PermissionId = permissionId
                    }).ToList();

                    await _context.AccountPermissions.AddRangeAsync(newPermissions);
                }

                await _context.SaveChangesAsync();

                return Json(new { success = true, message = "Cập nhật thành công!" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Đã xảy ra lỗi: {ex.Message}" });
            }
        }
    }
}
