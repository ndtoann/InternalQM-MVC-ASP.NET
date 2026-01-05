using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web_QM.Models;

namespace Web_QM.Areas.HR.Controllers
{
    [Area("HR")]
    public class MachineGroupController : Controller
    {
        private readonly QMContext _context;

        public MachineGroupController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ViewMachineGroup")]
        public async Task<IActionResult> Index()
        {
            var res = await _context.MachineGroups.AsNoTracking().ToListAsync();
            return View(res);
        }

        [Authorize(Policy = "AddMachineGroup")]
        public async Task<IActionResult> Add()
        {
            return View();
        }

        [Authorize(Policy = "AddMachineGroup")]
        [HttpPost]
        public async Task<IActionResult> Add(MachineGroup machineGroup)
        {
            if (!ModelState.IsValid)
            {
                return View(machineGroup);
            }
            if (await IsMachineTypeDuplicate(machineGroup.MachineType, machineGroup.GroupName))
            {
                TempData["ErrorMessage"] = "Nhóm máy đã tồn tại!";
                return View(machineGroup);
            }
            try
            {
                _context.MachineGroups.Add(machineGroup);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!";
                return View(machineGroup);
            }
        }

        [Authorize(Policy = "EditMachineGroup")]
        public async Task<IActionResult> Edit(long id)
        {
            if(id == null)
            {
                return NotFound();
            }
            var mgToEdit = await _context.MachineGroups.FindAsync(id);
            if (mgToEdit == null)
            {
                return NotFound();
            }
            return View(mgToEdit);
        }

        [Authorize(Policy = "EditMachineGroup")]
        [HttpPost]
        public async Task<IActionResult> Edit(long id, MachineGroup machineGroup)
        {
            if(id != machineGroup.Id)
            {
                return NotFound();
            }
            if (!ModelState.IsValid)
            {
                return View(machineGroup);
            }
            var mgToEdit = await _context.MachineGroups.AsNoTracking().FirstOrDefaultAsync(x => x.Id == machineGroup.Id);
            if (mgToEdit == null)
            {
                return NotFound();
            }
            if (await IsMachineTypeDuplicate(machineGroup.MachineType, machineGroup.GroupName, machineGroup.Id))
            {
                TempData["ErrorMessage"] = "Nhóm máy đã tồn tại!";
                return View(machineGroup);
            }
            try
            {
                _context.MachineGroups.Update(machineGroup);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Dữ liệu đã được lưu thành công!";
                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Đã xảy ra lỗi khi lưu dữ liệu. Vui lòng thử lại!";
                return View(machineGroup);
            }
        }

        [Authorize(Policy = "DeleteMachineGroup")]
        public async Task<IActionResult> Delete(long id)
        {
            if(id == null)
            {
                return NotFound();
            }
            var mgToDelete = await _context.MachineGroups.FirstOrDefaultAsync(x => x.Id == id);
            if (mgToDelete == null)
            {
                return NotFound();
            }
            try
            {
                await _context.Machine_MG
                     .Where(a => a.MachineGroupId == mgToDelete.Id)
                     .ExecuteDeleteAsync();

                _context.MachineGroups.Remove(mgToDelete);
                await _context.SaveChangesAsync();
                TempData["SuccessMessage"] = "Đã xóa thành công!";
            }
            catch (DbUpdateConcurrencyException)
            {
                TempData["ErrorMessage"] = "Xảy ra lỗi. Vui lòng thử lại sau!";
            }
            return RedirectToAction(nameof(Index));
        }

        [Authorize(Policy = "ViewMachineGroup")]
        public async Task<IActionResult> GetMachinesByGroup(long machineGroupId)
        {
            var machines = await _context.Machine_MG
                                         .Where(m => m.MachineGroupId == machineGroupId)
                                         .Select(m => new { m.MachineCode, m.Material })
                                         .ToListAsync();

            return Ok(machines);
        }

        [Authorize(Policy = "UpdateGroupMachine")]
        [HttpPost]
        public async Task<IActionResult> AddMachineToGroup([FromBody] Machine_MG model)
        {
            if (ModelState.IsValid)
            {
                var existingMachine = await _context.Machine_MG
                                                    .FirstOrDefaultAsync(m => m.MachineCode == model.MachineCode && m.MachineGroupId == model.MachineGroupId
                                                    && m.Material == model.Material);

                if (existingMachine != null)
                {
                    return Json(new { success = false, message = "Mã máy này đã tồn tại trong nhóm." });
                }

                _context.Machine_MG.Add(model);
                await _context.SaveChangesAsync();
                return Json(new { success = true, message = "Thêm máy thành công." });
            }
            return Json(new { success = false, message = "Dữ liệu không hợp lệ." });
        }

        [Authorize(Policy = "UpdateGroupMachine")]
        public async Task<IActionResult> RemoveMachineFromGroup([FromBody] Machine_MG model)
        {
            if (model == null)
            {
                return Json(new { success = false, message = "Dữ liệu không hợp lệ." });
            }
            var itemToDelete = await _context.Machine_MG
                .FirstOrDefaultAsync(m => m.MachineCode == model.MachineCode &&
                                          m.MachineGroupId == model.MachineGroupId &&
                                          m.Material == model.Material);
            if (itemToDelete == null)
            {
                return Json(new { success = false, message = "Không tìm thấy liên kết máy và vật liệu trong nhóm." });
            }
            _context.Machine_MG.Remove(itemToDelete);
            await _context.SaveChangesAsync();

            return Json(new { success = true });
        }

        [Authorize(Policy = "UpdateGroupMachine")]
        public async Task<IActionResult> GetMachineSelectOptions(string term)
        {
            var query = _context.Machines.AsQueryable();

            if (!string.IsNullOrEmpty(term))
            {
                string search = term.ToLower();
                query = query.Where(m => m.MachineCode.ToLower().Contains(search) || 
                                    m.MachineName.ToLower().Contains(search));
            }
            var machines = await query
                                 .Select(m => new
                                 {
                                     id = m.MachineCode,
                                     text = m.MachineCode + " - " + m.MachineName
                                 })
                                 .ToListAsync();
            return Ok(new { results = machines });
        }

        private async Task<bool> IsMachineTypeDuplicate(string machineType, string groupName, long? excludeId = null)
        {
            string normalizedType = machineType.ToLower();
            string normalizedGroup = groupName.ToLower();

            var query = _context.MachineGroups
                                .Where(mg => mg.MachineType.ToLower() == normalizedType || 
                                mg.GroupName.ToLower() == normalizedGroup);

            if (excludeId.HasValue)
            {
                query = query.Where(mg => mg.Id != excludeId.Value);
            }

            return await query.AnyAsync();
        }
    }
}
