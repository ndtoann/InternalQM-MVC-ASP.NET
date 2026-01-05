using DocumentFormat.OpenXml.InkML;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Web_QM.Models;

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class DashboardController : Controller
    {
        private readonly QMContext _context;

        public DashboardController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ViewAdmin")]
        public async Task<IActionResult> Index()
        {
            return View();
        }
    }
}
