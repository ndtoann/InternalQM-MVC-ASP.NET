using DocumentFormat.OpenXml.InkML;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Web_QM.Models;

namespace Web_QM.Areas.Production.Controllers
{
    [Area("Production")]
    public class DashboardController : Controller
    {
        private readonly QMContext _context;

        public DashboardController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ViewProduction")]
        public async Task<IActionResult> Index()
        {
            var cMachine = await _context.Machines.CountAsync();
            ViewBag.CountMachine = cMachine;

            return View();
        }

        [Authorize]
        public async Task<IActionResult> Logout()
        {
            await HttpContext.SignOutAsync(
            scheme: "SecurityScheme");

            HttpContext.Response.Cookies.Delete("email");

            return Redirect(nameof(Index));
        }
    }
}
