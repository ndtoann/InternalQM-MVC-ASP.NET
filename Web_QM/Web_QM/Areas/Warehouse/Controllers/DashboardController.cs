using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Web_QM.Areas.Demo.Controllers
{
    [Area("Warehouse")]
    public class DashboardController : Controller
    {
        [Authorize]
        public async Task<IActionResult> Index()
        {
            return View();
        }
    }
}
