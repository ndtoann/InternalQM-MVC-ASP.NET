using Microsoft.AspNetCore.Authentication;
using System.Security.Claims;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Common;
using Web_QM.Models;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Authorization;

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class AccountController : Controller
    {
        private readonly QMContext _context;

        public AccountController(QMContext context)
        {
            _context = context;
        }

        [Authorize]
        public async Task<IActionResult> Logout()
        {
            await HttpContext.SignOutAsync(
            scheme: "SecurityScheme");

            HttpContext.Response.Cookies.Delete("email");

            return Redirect(nameof(Index));
        }

        [Authorize(Policy = "ViewAccount")]
        public async Task<IActionResult> Index()
        {
            var res = await _context.Accounts.AsNoTracking().ToListAsync();
            return View(res);
        }

        [Authorize(Policy = "AddAccount")]
        public async Task<IActionResult> Add()
        {
            var employees = await _context.Employees.AsNoTracking().OrderBy(o => o.EmployeeCode).ToListAsync();
            var employeeList = employees.Select(e => new
            {
                Value = e.EmployeeCode,
                Text = e.EmployeeCode + " - " + e.EmployeeName
            });
            ViewBag.ListEmployee = new SelectList(employeeList, "Value", "Text");
            return View();
        }

        [Authorize(Policy = "AddAccount")]
        [HttpPost]
        public async Task<IActionResult> Add(Account account)
        {
            var employees = await _context.Employees.AsNoTracking().OrderBy(o => o.EmployeeCode).ToListAsync();
            var employeeList = employees.Select(e => new
            {
                Value = e.EmployeeCode,
                Text = e.EmployeeCode + " - " + e.EmployeeName
            });
            ViewBag.ListEmployee = new SelectList(employeeList, "Value", "Text");

            try
            {
                if (AccountExists(account.StaffCode))
                {
                    TempData["ErrorMessage"] = "Nhân viên đã có tài khoản!";
                    return View(account);
                }

                if (UsernameExists(account.UserName))
                {
                    TempData["ErrorMessage"] = "Tài khoản đã tồn tại!";
                    return View(account);
                }

                account.Salt ??= Guid.NewGuid().ToString();
                account.Password = account.UserName.ComputeSha256Hash(account.Salt, account.Password);

                var employee = await _context.Employees.FirstOrDefaultAsync(x => x.EmployeeCode == account.StaffCode);
                account.StaffName = employee.EmployeeName;

                account.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");

                _context.Accounts.Add(account);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Thêm tài khoản thành công!";
                return RedirectToAction("Index");
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
                return View(account);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
                return View(account);
            }
        }

        [Authorize(Policy = "EditAccount")]
        public async Task<IActionResult> Edit(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var accountToEdit = await _context.Accounts.AsNoTracking().FirstOrDefaultAsync(a => a.Id == id);
            if (accountToEdit == null)
            {
                return NotFound();
            }

            var userLogin = User.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Email)?.Value;
            if(userLogin == accountToEdit.UserName)
            {
                TempData["ErrorMessage"] = "Không thể sửa tài khoản bạn đang đăng nhập!";
                return RedirectToAction("Index");
            }
            return View(accountToEdit);
        }

        [Authorize(Policy = "EditAccount")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, Account account)
        {
            if(id != account.Id)
            {
                return NotFound();
            }

            if (!ModelState.IsValid)
            {
                return View(account);
            }
            var accountToEdit = await _context.Accounts.AsNoTracking().FirstOrDefaultAsync(a => a.Id == id);
            if (accountToEdit == null)
            {
                return NotFound();
            }

            var userLogin = User.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Email)?.Value;
            if (userLogin == accountToEdit.UserName)
            {
                TempData["ErrorMessage"] = "Không thể sửa tài khoản đang được đăng nhập!";
                return RedirectToAction("Index");
            }

            var existingAccount = await _context.Accounts.CountAsync(a => a.UserName == account.UserName && a.Id != account.Id);
            if (existingAccount > 0)
            {
                TempData["ErrorMessage"] = "Tài khoản đã tồn tại!";
                return View(account);
            }

            try
            {
                account.UpdatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Accounts.Update(account);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction("Index");
            }
            catch (DbUpdateConcurrencyException)
            {
                ModelState.AddModelError(string.Empty, "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!");
                return View(account);
            }
        }

        [Authorize(Policy = "EditAccount")]
        [HttpPost]
        public async Task<IActionResult> ChangePass(Account account, string newPass)
        {
            if (account == null || newPass ==null)
            {
                return NotFound();
            }

            try
            {
                account.Salt ??= Guid.NewGuid().ToString();
                account.Password = account.UserName.ComputeSha256Hash(account.Salt, newPass);
                account.UpdatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                
                _context.Accounts.Update(account);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Đổi mật khẩu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["SuccessMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
                return RedirectToAction(nameof(Index));
            }
        }

        [Authorize(Policy = "DeleteAccount")]
        public async Task<IActionResult> Delete(long id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var accountToDelete = await _context.Accounts.FindAsync(id);

            if (accountToDelete == null)
            {
                TempData["ErrorMessage"] = "Không tìm thấy tài khoản này!";
                return RedirectToAction(nameof(Index));
            }

            var accountIsLogin = User.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Email)?.Value;
            if (accountToDelete.UserName == accountIsLogin)
            {
                TempData["ErrorMessage"] = "Không thể xóa tài khoản do bạn đang đăng nhập!";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                await _context.AccountPermissions
                     .Where(a => a.AccountId == accountToDelete.Id)
                     .ExecuteDeleteAsync();

                _context.Accounts.Remove(accountToDelete);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Đã xóa tài khoản!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }

            return RedirectToAction(nameof(Index));
        }

        private bool AccountExists(string code)
        {
            return _context.Accounts.Any(e => e.StaffCode == code);
        }

        private bool UsernameExists(string username)
        {
            return _context.Accounts.Any(e => e.UserName == username);
        }
    }
}
