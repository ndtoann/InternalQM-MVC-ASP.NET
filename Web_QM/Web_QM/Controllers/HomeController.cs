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

        public IActionResult Login(string returnUrl = null)
        {
            ViewData["ReturnUrl"] = returnUrl;
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Login(string username, string password, bool rememberMe, string returnUrl = null)
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

            if (!string.IsNullOrEmpty(returnUrl) && Url.IsLocalUrl(returnUrl))
            {
                return LocalRedirect(returnUrl);
            }

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
        public async Task<IActionResult> SaveOpinion(string Type, string Title, string Content, IFormFile ImageFile)
        {
            try
            {
                if (ImageFile != null && ImageFile.Length > 0)
                {
                    var allowedExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif" };
                    var extension = Path.GetExtension(ImageFile.FileName).ToLower();

                    if (!allowedExtensions.Contains(extension) || ImageFile.Length > 5 * 1024 * 1024)
                    {
                        return Ok(new { success = false, message = "File không hợp lệ hoặc quá lớn" });
                    }
                }

                string fileName = null;
                if (ImageFile != null && ImageFile.Length > 0)
                {
                    string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/imgs/opinions");
                    if (!Directory.Exists(path)) Directory.CreateDirectory(path);

                    fileName = Guid.NewGuid().ToString() + Path.GetExtension(ImageFile.FileName);
                    string fullPath = Path.Combine(path, fileName);

                    using (var stream = new FileStream(fullPath, FileMode.Create))
                    {
                        await ImageFile.CopyToAsync(stream);
                    }
                }

                var userIdClaim = User.Claims.FirstOrDefault(c => c.Type == "EmployeeId")?.Value;

                var newOpinion = new Opinion
                {
                    Type = Type,
                    Title = Title,
                    Content = Content,
                    Img = fileName,
                    Status = 0,
                    CreatedBy = long.Parse(userIdClaim ?? "0"),
                    CreatedDate = DateOnly.FromDateTime(DateTime.Now)
                };

                _context.Opinions.Add(newOpinion);
                await _context.SaveChangesAsync();

                return Ok(new { success = true });
            }
            catch
            {
                return Ok(new { success = false });
            }
        }

        [Authorize]
        public async Task<IActionResult> GetTopSeniors()
        {
            var today = DateOnly.FromDateTime(DateTime.Now);

            var employees = _context.Employees
                .OrderBy(e => e.HireDate)
                .Take(10)
                .ToList()
                .Select(e => {
                    var totalDays = today.DayNumber - e.HireDate.DayNumber;
                    var years = totalDays / 365;
                    var months = (totalDays % 365) / 30;
                    var days = (totalDays % 365) % 30;

                    return new
                    {
                        e.EmployeeCode,
                        e.EmployeeName,
                        Avatar = string.IsNullOrEmpty(e.Avatar) ? "imgs/avatars/default.png" : "imgs/avatars/" + e.Avatar,
                        Tenure = $"{years} năm, {months} tháng, {days} ngày"
                    };
                });

            return Json(employees);
        }

        [Authorize]
        public async Task<IActionResult> GetMilestoneSeniors()
        {
            var today = DateOnly.FromDateTime(DateTime.Now);

            var tenYearsAgo = today.AddYears(-10);
            var elevenYearsAgo = today.AddYears(-11);

            var milestoneEmployees = await _context.Employees
                .Where(e => e.HireDate <= tenYearsAgo && e.HireDate > elevenYearsAgo)
                .Select(e => new {
                    e.EmployeeCode,
                    e.EmployeeName,
                    Avatar = string.IsNullOrEmpty(e.Avatar) ? "imgs/avatars/default.png" : "imgs/avatars/" + e.Avatar,
                    HireDate = e.HireDate.ToString("dd/MM/yyyy"),
                    TotalYears = today.Year - e.HireDate.Year,
                    Tenure = "10 Năm Cống Hiến"
                })
                .ToListAsync();

            return Ok(milestoneEmployees);
        }

        [Authorize]
        public IActionResult Denied()
        {
            return View();
        }
    }
}
