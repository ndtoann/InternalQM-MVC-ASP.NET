using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Memory;
using System.Diagnostics;
using System.Security.Claims;
using System.Text.RegularExpressions;
using Web_QM.Common;
using Web_QM.Models;

namespace Web_QM.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly QMContext _context;
        private readonly IMemoryCache _cache;

        private const int MaxFailedAttempts = 5;
        private readonly TimeSpan LockoutDuration = TimeSpan.FromMinutes(5);

        public HomeController(ILogger<HomeController> logger, QMContext context, IMemoryCache cache)
        {
            _logger = logger;
            _context = context;
            _cache = cache;
        }

        [Authorize]
        public IActionResult Index()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult Login()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Login(string username, string password, bool rememberMe)
        {
            string countKey = $"FailedLoginCount_{username}";
            string lockKey = $"LockoutTime_{username}";

            if (_cache.TryGetValue(lockKey, out DateTimeOffset lockedUntil) && lockedUntil > DateTimeOffset.UtcNow)
            {
                TimeSpan remainingTime = lockedUntil - DateTimeOffset.UtcNow;
                TempData["error"] = $"Tài khoản đã bị khóa do nhập sai quá nhiều lần. Vui lòng thử lại sau {remainingTime.Minutes} phút.";
                return View("Login");
            }

            _cache.Remove(lockKey);

            var acc = await _context.Accounts
                .Select(x => new
                {
                    x.Id,
                    x.UserName,
                    x.Password,
                    x.Salt
                })
                .FirstOrDefaultAsync(m => m.UserName == username);

            if (acc == null || !username.ValidPassword(acc.Salt, password, acc.Password))
            {
                int failedAttempts = _cache.TryGetValue(countKey, out int currentAttempts) ? currentAttempts : 0;
                failedAttempts++;

                if (failedAttempts >= MaxFailedAttempts)
                {
                    DateTimeOffset newLockoutTime = DateTimeOffset.UtcNow.Add(LockoutDuration);
                    _cache.Set(lockKey, newLockoutTime, newLockoutTime);

                    _cache.Remove(countKey);
                    TempData["error"] = $"Bạn đã nhập sai {MaxFailedAttempts} lần. Tài khoản bị khóa trong 5 phút.";
                }
                else
                {
                    _cache.Set(countKey, failedAttempts, TimeSpan.FromMinutes(10));
                    TempData["error"] = $"Tài khoản hoặc mật khẩu không chính xác!";
                }
                return View("Login");
            }

            var userBaseDetails = await (
                from account in _context.Accounts
                where account.Id == acc.Id
                join employee in _context.Employees on account.StaffCode equals employee.EmployeeCode into empGroup
                from employee in empGroup.DefaultIfEmpty()
                select new
                {
                    Account = new
                    {
                        account.UserName
                    },
                    Employee = new
                    {
                        employee.Id,
                        employee.EmployeeCode,
                        employee.EmployeeName,
                        employee.Department,
                        employee.Avatar
                    }
                }
            ).FirstOrDefaultAsync();

            if (userBaseDetails.Employee == null)
            {
                TempData["error"] = "Không tìm thấy thông tin nhân viên. Vui lòng liên hệ quản trị viên.";
                return View("Login");
            }

            var userPermissions = await (
                from rp in _context.AccountPermissions
                where rp.AccountId == acc.Id
                join permission in _context.Permissions on rp.PermissionId equals permission.Id
                select permission.ClaimValue
            ).ToListAsync();

            var claims = new List<Claim>
            {
                new Claim(ClaimTypes.Email, userBaseDetails.Account.UserName),
                new Claim("EmployeeId", userBaseDetails.Employee.Id.ToString()),
                new Claim("EmployeeCode", userBaseDetails.Employee.EmployeeCode),
                new Claim("EmployeeName", userBaseDetails.Employee.EmployeeName),
                new Claim("Department", userBaseDetails.Employee.Department),
                new Claim("Avatar", string.IsNullOrEmpty(userBaseDetails.Employee.Avatar) ? "default.png" : userBaseDetails.Employee.Avatar)
            };

            foreach (var claimValue in userPermissions)
            {
                claims.Add(new Claim("Permission", claimValue));
            }

            var identity = new ClaimsIdentity(claims, "SecurityScheme");
            var principal = new ClaimsPrincipal(identity);

            await HttpContext.SignInAsync(
                scheme: "SecurityScheme",
                principal: principal,
                properties: new AuthenticationProperties
                {
                    IsPersistent = rememberMe,
                    //ExpiresUtc = DateTimeOffset.UtcNow.AddMinutes(300)
                });

            _cache.Remove(countKey);
            _cache.Remove(lockKey);

            return RedirectToAction("index", "home");
        }

        [Authorize]
        public async Task<IActionResult> Logout()
        {
            await HttpContext.SignOutAsync(
            scheme: "SecurityScheme");

            HttpContext.Response.Cookies.Delete("email");

            return Redirect(nameof(Login));
        }

        [Authorize]
        public async Task<IActionResult> ResetPass()
        {
            return View();
        }

        [Authorize]
        [HttpPost]
        public async Task<IActionResult> ResetPass(string passOld, string passNew, string passNewRepeat)
        {
            if (passOld == null || passNew == null || passNewRepeat == null)
            {
                return Json(new { status = false, message = "Vui lòng nhập đầy đủ thông tin!" });
            }
            if (passNew.Length < 6 || !Regex.IsMatch(passNew, @"^[A-Za-z0-9@]*$"))
            {
                return Json(new { status = false, message = "Mật khẩu tối thiểu 6 ký tự và không chứa ký tự đặc biệt!" });
            }
            if (passNew != passNewRepeat)
            {
                return Json(new { status = false, message = "Mật khẩu mới không khớp!" });
            }
            var ussername = User.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Email)?.Value;
            if (ussername == null)
            {
                return Json(new { status = false, message = "Không tìm thấy thông tin người dùng!" });
            }
            var account = await _context.Accounts.FirstOrDefaultAsync(a => a.UserName == ussername);
            if (account == null)
            {
                return Json(new { status = false, message = "Tài khoản không tồn tại!" });
            }
            if (!ussername.ValidPassword(account.Salt, passOld, account.Password))
            {
                return Json(new { status = false, message = "Mật khẩu cũ không chính xác!" });
            }
            try
            {
                account.Password = account.UserName.ComputeSha256Hash(account.Salt, passNew);
                account.UpdatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Accounts.Update(account);
                await _context.SaveChangesAsync();
                return Json(new { status = true, message = "Đổi mật khẩu thành công!" });
            }
            catch (Exception ex)
            {
                return Json(new { status = false, message = "Đã xảy ra lỗi khi đổi mật khẩu. Vui lòng thử lại!" });
            }
        }

        [Authorize]
        [HttpPost]
        public async Task<IActionResult> SaveOpinion(Opinion model)
        {
            if (!ModelState.IsValid)
            {
                return Json(new { success = false, model = model });
            }
            try
            {
                model.Status = 0;
                model.CreatedBy = long.Parse(User.Claims.FirstOrDefault(c => c.Type == "EmployeeId")?.Value ?? "0");
                model.CreatedDate = DateOnly.FromDateTime(DateTime.Now);
                _context.Opinions.Add(model);
                await _context.SaveChangesAsync();
                return Json(new { success = true, model = model });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, model = model });
            }
        }

        [Authorize]
        public IActionResult Denied()
        {
            return View();
        }
    }
}
