using DocumentFormat.OpenXml.Bibliography;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using System.Reflection.PortableExecutable;
using System.Text.Json;
using Web_QM.Models;
using Web_QM.Models.ViewModels;

namespace Web_QM.Controllers
{
    public class MachineController : Controller
    {
        private readonly QMContext _context;

        private readonly List<string> toleranceTypes = new List<string> {
            "Độ chính xác",
            "Độ đảo trục chính",
            "Dung sai vị trí",
            "Độ song song",
            "Độ vuông góc"
        };

        public MachineController(QMContext context)
        {
            _context = context;
        }

        [Authorize(Policy = "ClientViewMachine")]
        public async Task<IActionResult> Index()
        {
            var today = DateOnly.FromDateTime(DateTime.Today);

            var cMachine = await _context.Machines.CountAsync();
            var cMachineGroup = await _context.MachineGroups.CountAsync();
            var cRepair = await _context.EquipmentRepairHistories.CountAsync(c => 
                                    (c.CompletionDate == null));
            var cMaintenance = await _context.MachineMaintenances.CountAsync(c => c.DateMonth >= today);

            ViewBag.CountMachine = cMachine;
            ViewBag.CountMachineGroup = cMachineGroup;
            ViewBag.CountRepair = cRepair;
            ViewBag.CountMaintenance = cMaintenance;
            return View();
        }

        [Authorize(Policy = "ClientViewMachine")]
        public async Task<IActionResult> List()
        {
            var chartData = _context.Machines
            .Where(m => (m.Status == "Đang hoạt động") &&
                        (!string.IsNullOrEmpty(m.Department)))
            .GroupBy(m => m.Department)
            .Select(g => new
            {
                Department = g.Key,
                MachineCount = g.Count()
            })
            .OrderByDescending(d => d.MachineCount)
            .ToList();
            string jsonChartData = JsonSerializer.Serialize(chartData);
            ViewBag.ChartData = jsonChartData;

            var typeMachineChartData = await _context.Machines
            .Where(m => m.Status == "Đang hoạt động")
            .GroupBy(m => m.TypeMachine ?? "Khác")
            .Select(g => new
            {
                TypeMachine = g.Key,
                MachineCount = g.Count()
            })
            .OrderByDescending(x => x.MachineCount)
            .ToListAsync();
            ViewBag.TypeMachineChartData = JsonSerializer.Serialize(typeMachineChartData);

            var cMachine = await _context.Machines.CountAsync();
            var cMachine1 = await _context.Machines.CountAsync(m => m.Status == "Đang hoạt động");
            var cMachine2 = await _context.Machines.CountAsync(m => m.Status == "Ngưng hoạt động");
            var cMachine3 = await _context.Machines.CountAsync(m => m.Status == "Đã bán");

            ViewBag.CountMachine = cMachine;
            ViewBag.CountMachineActive = cMachine1;
            ViewBag.CountMachineInActive = cMachine2;
            ViewBag.CountMachineSold = cMachine3;

            var departments = await _context.Departments.AsNoTracking().ToListAsync();
            var departmentsList = departments.Select(d => new SelectListItem
            {
                Value = d.DepartmentName,
                Text = d.DepartmentName,
            }).ToList();
            ViewData["Departments"] = departmentsList;

            return View();
        }

        [Authorize(Policy = "ClientViewMachine")]
        public async Task<IActionResult> GetMachines(string keyword, string department, string status)
        {
            var data = await _context.Machines
                .Where(m =>
                    string.IsNullOrEmpty(keyword) ||
                    (
                        m.MachineCode.ToLower().Contains(keyword.Trim().ToLower()) ||
                        m.MachineName.ToLower().Contains(keyword.Trim().ToLower())
                    )
                )
                .Where(m =>
                    string.IsNullOrEmpty(department) ||
                    m.Department == department
                )
                .Where(m =>
                    string.IsNullOrEmpty(status) ||
                    m.Status == status
                )
                .ToListAsync();

            return Json(new { data = data });
        }

        [Authorize(Policy = "ClientViewMachine")]
        public async Task<IActionResult> Detail(long id)
        {
            if(id == null)
            {
                return NotFound();
            }
            var machine = await _context.Machines.FindAsync(id);
            if (machine == null)
            {
                return NotFound();
            }

            var machineTolerance = new List<MachineToleranceDto>();

            var machineParameters = await _context.MachineParameters
                .Where(p => p.MachineCode == machine.MachineCode)
                .ToListAsync();

            foreach (var type in toleranceTypes)
            {
                var latestParam = machineParameters
                    .Where(p => p.Type == type)
                    .OrderByDescending(p => p.DateMonth)
                    .FirstOrDefault();

                machineTolerance.Add(new MachineToleranceDto
                {
                    Type = type,
                    LatestParameter = latestParam?.Parameters,
                    LatestDate = latestParam?.DateMonth,
                    LatestId = latestParam?.Id
                });
            }
            ViewBag.MachineTolerance = machineTolerance;

            return View(machine);
        }

        [Authorize(Policy = "ClientViewMachine")]
        public async Task<IActionResult> GetToleranceHistoryByType(string machineCode, string type)
        {
            if (string.IsNullOrEmpty(machineCode) || string.IsNullOrEmpty(type))
            {
                return BadRequest(new { success = false, message = "Thiếu mã máy hoặc loại dung sai." });
            }

            try
            {
                var history = await _context.MachineParameters
                    .Where(p => p.MachineCode == machineCode && p.Type == type)
                    .OrderByDescending(p => p.DateMonth)
                    .Select(p => new {
                        p.Id,
                        p.Type,
                        p.Parameters,
                        p.DateMonth
                    })
                    .ToListAsync();

                return Json(new { success = true, data = history });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = "Lỗi tải lịch sử dung sai." });
            }
        }

        [Authorize(Policy = "ClientViewRepair")]
        public async Task<IActionResult> GetEquipmentRepair(string? machineCode)
        {
            if(machineCode == null)
            {
                var list = await _context.EquipmentRepairHistories.AsNoTracking().OrderByDescending(r => r.DateMonth).ToListAsync();
                return Json(list);
            }
            var res = await _context.EquipmentRepairHistories.AsNoTracking().Where(r => r.EquipmentCode == machineCode).OrderByDescending(r => r.DateMonth).ToListAsync();
            return Json(res);
        }

        [Authorize(Policy = "ClientViewMaintenance")]
        public async Task<IActionResult> GetMachineMaintenance(string? machineCode)
        {
            if (machineCode == null)
            {
                var list = await _context.MachineMaintenances.AsNoTracking().OrderByDescending(r => r.DateMonth).ToListAsync();
                return Json(list);
            }
            var res = await _context.MachineMaintenances.AsNoTracking().Where(m => m.MachineCode == machineCode).OrderByDescending(m => m.DateMonth).ToListAsync();
            return Json(res);
        }

        [Authorize(Policy = "ClientViewMachineGroup")]
        public async Task<IActionResult> Group()
        {
            var allGroups = await _context.MachineGroups.ToListAsync();
            var allMachine_MGs = await _context.Machine_MG.ToListAsync();

            var treeView = BuildMachineTree(allGroups, allMachine_MGs);
            return View(treeView);
        }

        private MachineTreeView BuildMachineTree(List<MachineGroup> allGroups, List<Machine_MG> allMachine_MGs)
        {
            var treeData =
                from g in allGroups 
                join mg in allMachine_MGs on g.Id equals mg.MachineGroupId into machineLinks
                select new MachineGroupTreeNode
                {
                    MachineGroupId = g.Id,
                    GroupName = g.GroupName,
                    MachineType = g.MachineType,

                    Materials = (from link in machineLinks
                                 group link by link.Material into materialGroup
                                 select new MaterialNode
                                 {
                                     Material = materialGroup.Key,
                                     MachineCodes = materialGroup.Select(m => new MachineCodeNode
                                     {
                                         MachineCode = m.MachineCode
                                     }).ToList()
                                 })
                                 .OrderBy(m => m.Material)
                                 .ToList()
                };

            var treeView = new MachineTreeView();
            treeView.Groups.AddRange(treeData.OrderBy(g => g.GroupName).ToList());

            return treeView;
        }

        [Authorize(Policy = "ClientViewMachine")]
        public async Task<IActionResult> GetMachine(string machineCode)
        {
            var res = await _context.Machines.AsNoTracking().FirstOrDefaultAsync(m => m.MachineCode == machineCode);
            return Json(res);
        }

        [Authorize(Policy = "ClientViewRepair")]
        public async Task<IActionResult> Repair()
        {
            ViewBag.Departments = await _context.Departments
                                                .Select(d => new { d.DepartmentName })
                                                .ToListAsync();
            return View();
        }

        [Authorize(Policy = "ClientViewRepair")]
        [HttpGet]
        public async Task<IActionResult> GetEquipmentRepair(string department, string key, DateTime? startDate = null, DateTime? endDate = null)
        {
            try
            {
                var query = _context.EquipmentRepairHistories.AsNoTracking().AsQueryable();
                DateOnly? startOnly = startDate.HasValue ? DateOnly.FromDateTime(startDate.Value.Date) : (DateOnly?)null;
                DateOnly? endOnly = endDate.HasValue ? DateOnly.FromDateTime(endDate.Value.Date) : (DateOnly?)null;
                if (!string.IsNullOrEmpty(department))
                {
                    query = query.Where(r => r.Department == department);
                }
                if (!string.IsNullOrEmpty(key))
                {
                    string searchKey = key.ToLower();
                    query = query.Where(r =>
                        (r.EquipmentCode != null && r.EquipmentCode.ToLower().Contains(searchKey)) ||
                        (r.EquipmentName != null && r.EquipmentName.ToLower().Contains(searchKey))
                    );
                }
                query = query.Where(r =>
                    (!startOnly.HasValue || (r.DateMonth >= startOnly.Value)) &&
                    (!endOnly.HasValue || (r.DateMonth <= endOnly.Value))
                );

                var repairs = await query
                    .OrderByDescending(r => r.DateMonth)
                    .Take(1000)
                    .Select(r => new
                    {
                        Id = r.Id,
                        EquipmentName = r.EquipmentName,
                        EquipmentCode = r.EquipmentCode,
                        Department = r.Department,
                        Qty = r.Qty,
                        ErrorCondition = r.ErrorCondition,
                        Reason = r.Reason,
                        ProcessingMethod = r.ProcessingMethod,
                        DateMonth = r.DateMonth,
                        ConfirmCompletionDate = r.ConfirmCompletionDate,
                        CompletionDate = r.CompletionDate,
                        RepairCosts = r.RepairCosts,
                        Note = r.Note
                    })
                    .ToListAsync();
                return Json(new { success = true, data = repairs });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Lỗi Server: {ex.Message}", details = ex.InnerException?.Message });
            }
        }

        [Authorize(Policy = "ClientViewRepair")]
        public async Task<IActionResult> GetReplacementsByRepairId(long repairId)
        {
            if (repairId == null)
            {
                return BadRequest(new { success = false, message = "ID không hợp lệ." });
            }

            try
            {
                var replacements = await _context.ReplacementEquipmentAndSupplies
                    .Where(r => r.EquipmentRepairId == repairId)
                    .Select(r => new
                    {
                        Id = r.Id,
                        EquipmentAndlSupplies = r.EquipmentAndlSupplies,
                        FilePdf = r.FilePdf
                    })
                    .ToListAsync();
                return Json(new { success = true, data = replacements });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = "Lỗi máy chủ khi lấy dữ liệu!", error = ex.Message });
            }
        }

        [Authorize(Policy = "ClientViewMaintenance")]
        public async Task<IActionResult> ScheduleMaintenance()
        {
            var allMachineCodes = await _context.Machines
                .Select(m => m.MachineCode)
                .Distinct()
                .OrderBy(c => c)
                .ToListAsync();

            var availableYears = await _context.MachineMaintenances
                .Select(m => m.DateMonth.Year)
                .Distinct()
                .OrderByDescending(y => y)
                .ToListAsync();

            int defaultYear = availableYears.Any() ? availableYears.Max() : DateTime.Today.Year;

            ViewBag.AllMachineCodes = allMachineCodes;
            ViewBag.AvailableYears = availableYears;
            ViewBag.DefaultYear = defaultYear;

            return View();
        }

        [Authorize(Policy = "ClientViewMaintenance")]
        public IActionResult GetMaintenanceData(int year)
        {
            var yearSchedules = _context.MachineMaintenances
                .Where(m => m.DateMonth.Year == year)
                .ToList();

            var scheduleMap = yearSchedules
                .GroupBy(m => m.MachineCode)
                .ToDictionary(
                    g => g.Key,
                    g => g.GroupBy(m => m.DateMonth.Month)
                          .ToDictionary(
                              mGroup => mGroup.Key,
                              dGroup => dGroup.GroupBy(d => d.DateMonth.Day)
                                              .ToDictionary(
                                                  dayGroup => dayGroup.Key,
                                                  dayGroup => dayGroup.ToList()
                                              )
                          )
                );

            return Json(scheduleMap);
        }
    }
}