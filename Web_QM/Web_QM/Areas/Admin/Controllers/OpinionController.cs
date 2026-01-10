using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    [Authorize(Policy = "ViewOpinion")]
    public class OpinionController : Controller
    {
        private readonly QMContext _context;

        public OpinionController(QMContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> Index()
        {
            return View();
        }

        [HttpGet]
        public async Task<IActionResult> GetList(string keyword, int? status)
        {
            var query = from o in _context.Opinions
                        join e in _context.Employees on o.CreatedBy equals e.Id into joined
                        from e in joined.DefaultIfEmpty()
                        select new
                        {
                            o.Id,
                            o.Title,
                            o.Type,
                            o.Content,
                            o.Img,
                            Status = o.Status.HasValue ? (int)o.Status.Value : 0,
                            o.CreatedDate,
                            EmployeeCode = e != null ? e.EmployeeCode : "N/A",
                            EmployeeName = e != null ? e.EmployeeName : "N/A"
                        };

            if (!string.IsNullOrEmpty(keyword))
            {
                query = query.Where(x => x.Title.Contains(keyword) || x.Content.Contains(keyword));
            }

            if (status.HasValue)
            {
                query = query.Where(x => x.Status == status.Value);
            }

            var data = await query.OrderByDescending(x => x.CreatedDate)
                                  .ThenByDescending(x => x.Id)
                                  .ToListAsync();

            return Json(data);
        }

        public async Task<IActionResult> UpdateStatus(long id)
        {
            var item = await _context.Opinions.FindAsync(id);
            if (item == null) return Json(new { success = false });

            try
            {
                item.Status = 1;
                await _context.SaveChangesAsync();
                return Json(new { success = true });
            }
            catch (DbUpdateConcurrencyException)
            {
                return Json(new { success = false });
            }
        }

        public async Task<IActionResult> Delete(long id)
        {
            var item = await _context.Opinions.FindAsync(id);
            if (item == null) return Json(new { success = false });

            try
            {
                if (item.Img != null)
                {
                    string oldFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "imgs", "opinions", item.Img);
                    if (System.IO.File.Exists(oldFilePath))
                    {
                        System.IO.File.Delete(oldFilePath);
                    }
                }

                _context.Opinions.Remove(item);
                await _context.SaveChangesAsync();
                return Json(new { success = true });
            }
            catch (DbUpdateConcurrencyException)
            {
                return Json(new { success = false });
            }
        }
    }
}
