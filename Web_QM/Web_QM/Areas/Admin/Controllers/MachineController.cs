using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;

namespace Web_QM.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class MachineController : Controller
    {
        private readonly QMContext _context;
        private readonly IWebHostEnvironment _env;

        public MachineController(QMContext context, IWebHostEnvironment env)
        {
            _context = context;
            _env = env;
        }

        [Authorize(Policy = "ViewMachine")]
        public async Task<IActionResult> Index(string status, string department, string key)
        {
            var res = await _context.Machines.AsNoTracking()
                                            .Where(m =>
                                                (string.IsNullOrEmpty(status) || m.Status == status)
                                                &&
                                                (string.IsNullOrEmpty(department) || m.Department == department)
                                                &&
                                                (string.IsNullOrEmpty(key) || m.MachineCode.ToLower().Contains(key.ToLower()) ||
                                                m.MachineName.ToLower().Contains(key.ToLower()))
                                            )
                                            .ToListAsync();
            var statusOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "", Text = "Trạng thái máy" },
                new SelectListItem { Value = "Đang hoạt động", Text = "Đang hoạt động" },
                new SelectListItem { Value = "Ngưng hoạt động", Text = "Ngưng hoạt động" },
                new SelectListItem { Value = "Đã bán", Text = "Đã bán" }
            };
            ViewData["Status"] = new SelectList(statusOptions, "Value", "Text", status);

            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName,
                Selected = d.DepartmentName == department
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewData["Departments"] = departmentsList;
            ViewBag.KeySearch = key;
            return View(res);
        }

        [Authorize(Policy = "AddMachine")]
        public async Task<IActionResult> Add()
        {
            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName,
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewData["Departments"] = departmentsList;
            return View();
        }

        [Authorize(Policy = "AddMachine")]
        [HttpPost]
        public async Task<IActionResult> Add(Machine machine, IFormFile pictureFile)
        {
            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewData["Departments"] = departmentsList;
            var existingMachine = await _context.Machines.AnyAsync(m => m.MachineCode == machine.MachineCode);
            if (existingMachine)
            {
                TempData["ErrorMessage"] = "Mã máy đã tồn tại!";
                return View(machine);
            }
            try
            {
                string picture = UploadPicture(pictureFile, machine.MachineCode);
                machine.Picture = picture;
                machine.CreatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Machines.Add(machine);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!";
                return View(machine);
            }
        }

        [Authorize(Policy = "EditMachine")]
        public async Task<IActionResult> Edit(long id)
        {
            if(id == null)
            {
                return NotFound();
            }
            var machineToEdit = await _context.Machines.FindAsync(id);
            if(machineToEdit == null)
            {
                return NotFound();
            }
            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewData["Departments"] = departmentsList;
            return View(machineToEdit);
        }

        [Authorize(Policy = "EditMachine")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, Machine machine, IFormFile pictureFile)
        {
            if(id != machine.Id)
            {
                return NotFound();
            }
            var machineToEdit = await _context.Machines.AsNoTracking().FirstOrDefaultAsync(m => m.Id == machine.Id);
            if(machineToEdit == null)
            {
                return NotFound();
            }
            var departments = _context.Departments.ToList();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName
            }).ToList();
            departmentsList.Insert(0, new SelectListItem { Value = "", Text = "Chọn bộ phận" });
            ViewData["Departments"] = departmentsList;

            var existingMachine = await _context.Machines.AnyAsync(m => m.MachineCode == machine.MachineCode && m.Id != machine.Id);
            if (existingMachine)
            {
                TempData["ErrorMessage"] = "Mã máy đã tồn tại!";
                return View(machine);
            }
            try
            {
                if (pictureFile != null && pictureFile.Length > 0)
                {
                    if (!string.IsNullOrEmpty(machineToEdit.Picture))
                    {
                        string oldFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "imgs", "machines", machineToEdit.Picture);
                        if (System.IO.File.Exists(oldFilePath))
                        {
                            System.IO.File.Delete(oldFilePath);
                        }
                    }

                    string newAvatarName = UploadPicture(pictureFile, machine.MachineCode);
                    machine.Picture = newAvatarName;
                }
                else
                {
                    machine.Picture = machineToEdit.Picture;
                }
                machine.CreatedDate = machineToEdit.CreatedDate;
                machine.UpdatedDate = DateTime.Now.ToString("HH:mm:ss-dd/MM/yyyy");
                _context.Machines.Update(machine);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!";
                return View(machine);
            }
        }

        [Authorize(Policy = "DeleteMachine")]
        public async Task<IActionResult> Delete(long id)
        {
            if(id == null)
            {
                return NotFound();
            }
            var machineToDel = await _context.Machines.FindAsync(id);
            if (machineToDel == null)
            {
                return NotFound();
            }
            try
            {
                await _context.Machine_MG
                     .Where(a => a.MachineCode == machineToDel.MachineCode)
                     .ExecuteDeleteAsync();

                await _context.MachineParameters
                     .Where(a => a.MachineCode == machineToDel.MachineCode)
                     .ExecuteDeleteAsync();

                await _context.MachineMaintenances
                     .Where(a => a.MachineCode == machineToDel.MachineCode)
                     .ExecuteDeleteAsync();

                if (machineToDel.Picture != null)
                {
                    string oldFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "imgs", "machines", machineToDel.Picture);
                    if (System.IO.File.Exists(oldFilePath))
                    {
                        System.IO.File.Delete(oldFilePath);
                    }
                }

                _context.Machines.Remove(machineToDel);
                _context.SaveChanges();
                TempData["SuccessMessage"] = "Đã xóa toàn bộ thông tin của máy!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }
            return RedirectToAction(nameof(Index));
        }

        [Authorize(Policy = "ViewMachine")]
        public async Task<IActionResult> GetMachineParametersByCode(string machineCode)
        {
            if (string.IsNullOrEmpty(machineCode))
            {
                return BadRequest(new { success = false, message = "MachineCode không được để trống." });
            }

            try
            {
                var parameters = await _context.MachineParameters
                    .Where(p => p.MachineCode == machineCode)
                    .OrderByDescending(p => p.DateMonth)
                    .ToListAsync();

                return Json(parameters);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = "Lỗi server khi tải dữ liệu dung sai.", error = ex.Message });
            }
        }

        [Authorize(Policy = "EditMachine")]
        [HttpPost]
        public async Task<IActionResult> AddMachineParameter([FromBody] MachineParameter newParameter)
        {
            if (newParameter.Parameters <= 0 || newParameter.Parameters >= 1000)
            {
                return Json(new
                {
                    success = false,
                    message = "Thông số dung sai phải lớn hơn 0!"
                });
            }
            var isDuplicate = await _context.MachineParameters
                .AnyAsync(p =>
                    p.MachineCode == newParameter.MachineCode &&
                    p.Type == newParameter.Type &&
                    p.DateMonth == newParameter.DateMonth);

            if (isDuplicate)
            {
                string dateDisplay = newParameter.DateMonth.ToString("dd/MM/yyyy");
                return Json(new
                {
                    success = false,
                    message = $"Dung sai loại '{newParameter.Type}' đã tồn tại cho ngày {dateDisplay}. Vui lòng kiểm tra lại."
                });
            }
            try
            {
                _context.MachineParameters.Add(newParameter);
                await _context.SaveChangesAsync();

                return Json(new { success = true, message = "Thêm dung sai thành công!", data = newParameter });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = "Lỗi server khi thêm dung sai.", error = ex.Message });
            }
        }

        [Authorize(Policy = "EditMachine")]
        public async Task<IActionResult> DeleteMachineParameter(long id)
        {
            if (id <= 0)
            {
                return Json(new { success = false, message = "ID dung sai không hợp lệ." });
            }

            try
            {
                var parameter = await _context.MachineParameters.FindAsync(id);

                if (parameter == null)
                {
                    return Json(new { success = false, message = "Không tìm thấy dung sai cần xóa." });
                }

                _context.MachineParameters.Remove(parameter);
                await _context.SaveChangesAsync();

                return Json(new { success = true, message = "Xóa dung sai thành công!" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = "Lỗi server khi xóa dung sai.", error = ex.Message });
            }
        }

        [Authorize(Policy = "ViewMachine")]
        public async Task<IActionResult> GetMachineGroupsByCode(string machineCode)
        {
            if (string.IsNullOrEmpty(machineCode))
            {
                return BadRequest(new { message = "Mã máy không hợp lệ." });
            }
            var machineGroups = await _context.Machine_MG
                .Where(mg => mg.MachineCode == machineCode)
                .Join(
                    _context.MachineGroups,
                    machineMG => machineMG.MachineGroupId,
                    group => group.Id,
                    (machineMG, group) => new
                    {
                        Id = group.Id,
                        GroupName = group.GroupName,
                        MachineType = group.MachineType,
                        Material = machineMG.Material,
                        Standard = group.Standard
                    }
                )
                .ToListAsync();

            if (machineGroups == null || !machineGroups.Any())
            {
                return Ok(new List<object>());
            }
            return Ok(machineGroups);
        }

        private string UploadPicture(IFormFile file, string machineCode)
        {
            if (file == null || file.Length == 0)
            {
                return null;
            }

            var allowedExtensions = new[] { ".png", ".jpeg", ".jpg" };
            const int maxFileSizeMB = 5;
            const long maxFileSizeInBytes = maxFileSizeMB * 1024 * 1024;
            if (file.Length > maxFileSizeInBytes)
            {
                Console.WriteLine($"Lỗi: Dung lượng file quá lớn. Tối đa là {maxFileSizeMB}MB.");
                return null;
            }

            try
            {
                string fileExtension = Path.GetExtension(file.FileName).ToLowerInvariant();

                if (!allowedExtensions.Contains(fileExtension))
                {
                    Console.WriteLine($"Lỗi: Định dạng file không được hỗ trợ. Chỉ cho phép các định dạng: {string.Join(", ", allowedExtensions)}.");
                    return null;
                }

                string newFileName = $"machine_{machineCode}{fileExtension}";
                string uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "imgs", "machines");
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }
                string filePath = Path.Combine(uploadPath, newFileName);
                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(fileStream);
                }

                return newFileName;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi tải ảnh lên: {ex.Message}");
                return null;
            }
        }
    }
}
