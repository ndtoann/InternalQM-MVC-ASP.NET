using Microsoft.AspNetCore.Mvc;

namespace Web_QM.Areas.Demo.Controllers
{
    [Area("Demo")]
    public class DashboardController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
