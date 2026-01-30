using Microsoft.EntityFrameworkCore;
using System.Text.Json;
using Web_QM.Models;
using Microsoft.AspNetCore.Http;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddRazorPages();
builder.Services.AddHttpContextAccessor();

builder.Services.AddMemoryCache();

var connnStr = builder.Configuration.GetConnectionString("SqlServer");
builder.Services.AddDbContext<QMContext>(options =>
    options.UseSqlServer(connnStr, sqlServerOptionsAction: sqlOptions =>
    {
        sqlOptions.EnableRetryOnFailure(
            maxRetryCount: 10,
            maxRetryDelay: TimeSpan.FromSeconds(30),
            errorNumbersToAdd: null);
    }));

builder.Services.AddAuthentication("SecurityScheme")
    .AddCookie("SecurityScheme", options =>
    {
        options.Cookie = new CookieBuilder
        {
            HttpOnly = true,
            Name = ".aspNetCore.Web_QM.Cookie",
            Path = "/",
            SameSite = SameSiteMode.Strict,
            SecurePolicy = CookieSecurePolicy.SameAsRequest
        };
        options.ExpireTimeSpan = TimeSpan.FromMinutes(300);
        options.LoginPath = new PathString("/home/login");
        options.ReturnUrlParameter = "ReturnUrl";
        options.SlidingExpiration = true;
        options.AccessDeniedPath = "/home/denied";
    });

builder.Services.AddAuthorization(options =>
{
    //quản lý dashboard
    options.AddPolicy("ViewAdmin", policy =>
        policy.RequireClaim("Permission", "Admin.View"));

    options.AddPolicy("ViewHr", policy =>
        policy.RequireClaim("Permission", "Hr.View"));

    options.AddPolicy("ViewProduction", policy =>
        policy.RequireClaim("Permission", "Production.View"));

    options.AddPolicy("ViewWareHouse", policy =>
        policy.RequireClaim("Permission", "WareHouse.View"));

    //thông báo
    options.AddPolicy("ViewNotifi", policy =>
        policy.RequireClaim("Permission", "Notification.View"));

    options.AddPolicy("EditNotifi", policy =>
        policy.RequireClaim("Permission", "Notification.Edit"));

    options.AddPolicy("DeleteNotifi", policy =>
        policy.RequireClaim("Permission", "Notification.Delete"));

    //đóng góp, ý kiến
    options.AddPolicy("ViewOpinion", policy =>
        policy.RequireClaim("Permission", "Opinion.View"));

    //tài khoản
    options.AddPolicy("ViewAccount", policy =>
        policy.RequireClaim("Permission", "Account.View"));

    options.AddPolicy("AddAccount", policy =>
        policy.RequireClaim("Permission", "Account.Add"));

    options.AddPolicy("EditAccount", policy =>
        policy.RequireClaim("Permission", "Account.Edit"));

    options.AddPolicy("DeleteAccount", policy =>
        policy.RequireClaim("Permission", "Account.Delete"));

    //các quyền nhân viên
    options.AddPolicy("ViewPermission", policy =>
       policy.RequireClaim("Permission", "Permission.View"));

    options.AddPolicy("AddPermission", policy =>
       policy.RequireClaim("Permission", "Permission.Add"));

    options.AddPolicy("EditPermission", policy =>
       policy.RequireClaim("Permission", "Permission.Edit"));

    options.AddPolicy("DeletePermission", policy =>
       policy.RequireClaim("Permission", "Permission.Delete"));

    options.AddPolicy("EmplPermission", policy =>
       policy.RequireClaim("Permission", "EmplPermission.Update"));

    //bộ phận
    options.AddPolicy("ViewDepartment", policy =>
       policy.RequireClaim("Permission", "Department.View"));

    options.AddPolicy("AddDepartment", policy =>
       policy.RequireClaim("Permission", "Department.Add"));

    options.AddPolicy("EditDepartment", policy =>
       policy.RequireClaim("Permission", "Department.Edit"));

    options.AddPolicy("DeleteDepartment", policy =>
       policy.RequireClaim("Permission", "Department.Delete"));

    //nhân viên
    options.AddPolicy("ViewEmployee", policy =>
        policy.RequireClaim("Permission", "ManageEmployee.View"));

    options.AddPolicy("AddEmployee", policy =>
        policy.RequireClaim("Permission", "ManageEmployee.Add"));

    options.AddPolicy("EditEmployee", policy =>
        policy.RequireClaim("Permission", "ManageEmployee.Edit"));

    options.AddPolicy("DeleteEmployee", policy =>
       policy.RequireClaim("Permission", "ManageEmployee.Delete"));

    options.AddPolicy("DataEmpl", policy =>
       policy.RequireClaim("Permission", "DataEmpl.Update"));

    options.AddPolicy("WorkHistoryEmpl", policy =>
       policy.RequireClaim("Permission", "WorkHistoryEmpl.Update"));

    options.AddPolicy("EmplViolation5S", policy =>
       policy.RequireClaim("Permission", "EmplViolation5S.Update"));

    //bài kiểm tra
    options.AddPolicy("ViewExam", policy =>
       policy.RequireClaim("Permission", "Exam.View"));

    options.AddPolicy("AddExam", policy =>
       policy.RequireClaim("Permission", "Exam.Add"));

    options.AddPolicy("EditExam", policy =>
       policy.RequireClaim("Permission", "Exam.Edit"));

    options.AddPolicy("DeleteExam", policy =>
       policy.RequireClaim("Permission", "Exam.Delete"));

    options.AddPolicy("UpdateQuestion", policy =>
       policy.RequireClaim("Permission", "Question.Update"));


    //kết quả kiểm tra
    options.AddPolicy("ViewResultExam", policy =>
       policy.RequireClaim("Permission", "ResultExam.View"));

    options.AddPolicy("EditResultExam", policy =>
       policy.RequireClaim("Permission", "ResultExam.Edit"));

    options.AddPolicy("DeleteResultExam", policy =>
       policy.RequireClaim("Permission", "ResultExam.Delete"));

    //thống kê lỗi sx
    options.AddPolicy("ViewProductionDefect", policy =>
       policy.RequireClaim("Permission", "ManageProductionDefect.View"));

    options.AddPolicy("AddProductionDefect", policy =>
       policy.RequireClaim("Permission", "ManageProductionDefect.Add"));

    options.AddPolicy("EditProductionDefect", policy =>
       policy.RequireClaim("Permission", "ManageProductionDefect.Edit"));

    options.AddPolicy("DeleteProductionDefect", policy =>
       policy.RequireClaim("Permission", "ManageProductionDefect.Delete"));

    //kaizen
    options.AddPolicy("ViewKaizen", policy =>
       policy.RequireClaim("Permission", "ManageKaizen.View"));

    options.AddPolicy("AddKaizen", policy =>
       policy.RequireClaim("Permission", "ManageKaizen.Add"));

    options.AddPolicy("EditKaizen", policy =>
       policy.RequireClaim("Permission", "ManageKaizen.Edit"));

    options.AddPolicy("DeleteKaizen", policy =>
       policy.RequireClaim("Permission", "ManageKaizen.Delete"));

    //năng suất
    options.AddPolicy("ViewProductivity", policy =>
       policy.RequireClaim("Permission", "ManageProductivity.View"));

    options.AddPolicy("AddProductivity", policy =>
       policy.RequireClaim("Permission", "ManageProductivity.Add"));

    options.AddPolicy("EditProductivity", policy =>
       policy.RequireClaim("Permission", "ManageProductivity.Edit"));

    options.AddPolicy("DeleteProductivity", policy =>
       policy.RequireClaim("Permission", "ManageProductivity.Delete"));

    //lỗi 5S
    options.AddPolicy("ViewViolation5S", policy =>
       policy.RequireClaim("Permission", "ManageViolation5S.View"));

    options.AddPolicy("AddViolation5S", policy =>
       policy.RequireClaim("Permission", "ManageViolation5S.Add"));

    options.AddPolicy("EditViolation5S", policy =>
       policy.RequireClaim("Permission", "ManageViolation5S.Edit"));

    options.AddPolicy("DeleteViolation5S", policy =>
       policy.RequireClaim("Permission", "ManageViolation5S.Delete"));

    //quản lý máy
    options.AddPolicy("ViewMachine", policy =>
       policy.RequireClaim("Permission", "ManageMachine.View"));

    options.AddPolicy("AddMachine", policy =>
       policy.RequireClaim("Permission", "ManageMachine.Add"));

    options.AddPolicy("EditMachine", policy =>
       policy.RequireClaim("Permission", "ManageMachine.Edit"));

    options.AddPolicy("DeleteMachine", policy =>
       policy.RequireClaim("Permission", "ManageMachine.Delete"));

    //nhóm máy
    options.AddPolicy("ViewMachineGroup", policy =>
       policy.RequireClaim("Permission", "ManageGroupMachine.View"));

    options.AddPolicy("AddMachineGroup", policy =>
       policy.RequireClaim("Permission", "ManageGroupMachine.Add"));

    options.AddPolicy("EditMachineGroup", policy =>
       policy.RequireClaim("Permission", "ManageGroupMachine.Edit"));

    options.AddPolicy("DeleteMachineGroup", policy =>
       policy.RequireClaim("Permission", "ManageGroupMachine.Delete"));

    options.AddPolicy("UpdateGroupMachine", policy =>
       policy.RequireClaim("Permission", "ManageGroupMachine.AddMachine"));

    //sữa chữa thiết bị máy móc
    options.AddPolicy("ViewMachineRepair", policy =>
       policy.RequireClaim("Permission", "ManageRepairMachine.View"));

    options.AddPolicy("AddMachineRepair", policy =>
       policy.RequireClaim("Permission", "ManageRepairMachine.Add"));

    options.AddPolicy("EditMachineRepair", policy =>
       policy.RequireClaim("Permission", "ManageRepairMachine.Edit"));

    options.AddPolicy("DeleteMachineRepair", policy =>
       policy.RequireClaim("Permission", "ManageRepairMachine.Delete"));

    //bảo dưỡng thiết bị máy móc
    options.AddPolicy("ViewMachineMaintenance", policy =>
       policy.RequireClaim("Permission", "ManageMaintenance.View"));

    options.AddPolicy("AddMachineMaintenance", policy =>
       policy.RequireClaim("Permission", "ManageMaintenance.Add"));

    options.AddPolicy("EditMachineMaintenance", policy =>
       policy.RequireClaim("Permission", "ManageMaintenance.Edit"));

    options.AddPolicy("DeleteMachineMaintenance", policy =>
       policy.RequireClaim("Permission", "ManageMaintenance.Delete"));

    //lương nhân viên
    options.AddPolicy("ViewSalary", policy =>
       policy.RequireClaim("Permission", "ManageSalary.View"));

    options.AddPolicy("AddSalary", policy =>
       policy.RequireClaim("Permission", "ManageSalary.Add"));

    options.AddPolicy("EditSalary", policy =>
       policy.RequireClaim("Permission", "ManageSalary.Edit"));

    options.AddPolicy("DeleteSalary", policy =>
       policy.RequireClaim("Permission", "ManageSalary.Delete"));

    options.AddPolicy("ViewMonthlyPayroll", policy =>
       policy.RequireClaim("Permission", "ManageMonthlySalary.View"));

    options.AddPolicy("UpdateMonthlyPayroll", policy =>
       policy.RequireClaim("Permission", "ManageMonthlySalary.Update"));

    //bảng chấm công
    options.AddPolicy("ViewTimeSheet", policy =>
       policy.RequireClaim("Permission", "TimeSheet.View"));

    options.AddPolicy("ApproveTimeSheet", policy =>
       policy.RequireClaim("Permission", "TimeSheet.Approve"));

    options.AddPolicy("ViewAllTimeSheet", policy =>
       policy.RequireClaim("Permission", "TimeSheet.ViewAll"));

    //đồ gá - dao cụ
    options.AddPolicy("ViewTool", policy =>
       policy.RequireClaim("Permission", "Tool.View"));

    options.AddPolicy("AddTool", policy =>
       policy.RequireClaim("Permission", "Tool.Add"));

    options.AddPolicy("EditTool", policy =>
       policy.RequireClaim("Permission", "Tool.Edit"));

    options.AddPolicy("DeleteTool", policy =>
       policy.RequireClaim("Permission", "Tool.Delete"));

    options.AddPolicy("ViewToolSupplyLog", policy =>
       policy.RequireClaim("Permission", "ToolSupplyLog.View"));

    options.AddPolicy("AddToolSupplyLog", policy =>
       policy.RequireClaim("Permission", "ToolSupplyLog.Add"));

    options.AddPolicy("EditToolSupplyLog", policy =>
       policy.RequireClaim("Permission", "ToolSupplyLog.Edit"));

    options.AddPolicy("DeleteToolSupplyLog", policy =>
       policy.RequireClaim("Permission", "ToolSupplyLog.Delete"));

    options.AddPolicy("ViewIssueReturnLog", policy =>
       policy.RequireClaim("Permission", "IssueReturnLog.View"));

    options.AddPolicy("AddIssueReturnLog", policy =>
       policy.RequireClaim("Permission", "IssueReturnLog.Add"));

    options.AddPolicy("EditIssueReturnLog", policy =>
       policy.RequireClaim("Permission", "IssueReturnLog.Edit"));

    options.AddPolicy("DeleteIssueReturnLog", policy =>
       policy.RequireClaim("Permission", "IssueReturnLog.Delete"));

    //quy trình sản xuất
    options.AddPolicy("ViewProductionProcessess", policy =>
       policy.RequireClaim("Permission", "ProductionProcessess.View"));

    options.AddPolicy("AddProductionProcessess", policy =>
       policy.RequireClaim("Permission", "ProductionProcessess.Add"));

    options.AddPolicy("EditProductionProcessess", policy =>
       policy.RequireClaim("Permission", "ProductionProcessess.Edit"));

    options.AddPolicy("DeleteProductionProcessess", policy =>
       policy.RequireClaim("Permission", "ProductionProcessess.Delete"));

    //client
    options.AddPolicy("ClientViewEmpl", policy =>
       policy.RequireClaim("Permission", "EmployeeSameDepartment.View"));

    options.AddPolicy("ClientAddFeedbackEmpl", policy =>
       policy.RequireClaim("Permission", "Employee.AddFeedbackEmpl"));

    options.AddPolicy("ClientViewMachine", policy =>
       policy.RequireClaim("Permission", "Machine.View"));

    options.AddPolicy("ClientViewMachineGroup", policy =>
       policy.RequireClaim("Permission", "MachineGroup.View"));

    options.AddPolicy("ClientViewRepair", policy =>
       policy.RequireClaim("Permission", "MachineRepair.View"));

    options.AddPolicy("ClientViewMaintenance", policy =>
       policy.RequireClaim("Permission", "MachineMaintenance.View"));

    options.AddPolicy("ClientViewProductivity", policy =>
       policy.RequireClaim("Permission", "Productivity.View"));

    options.AddPolicy("ClientViewViolation5S", policy =>
       policy.RequireClaim("Permission", "EmplViolation5S.View"));

    options.AddPolicy("ClientViewProductionDefect", policy =>
       policy.RequireClaim("Permission", "ProductionDefect.View"));
});

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.Use(async (context, next) =>
{
    await next();

    var now = DateTime.Now;
    var path = context.Request.Path;

    string[] extensionsToIgnore = { ".js", ".css", ".png", ".jpg", "jpeg",".gif", ".svg", ".ico", ".woff", ".woff2" };
    if (extensionsToIgnore.Any(ext => path.Value.EndsWith(ext, StringComparison.OrdinalIgnoreCase))) return;

    var statusCode = context.Response.StatusCode;
    var userName = context.User.FindFirst("EmployeeCode")?.Value
                   ?? context.User.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value
                   ?? "Anonymous";
    var method = context.Request.Method;
    var ip = context.Connection.RemoteIpAddress?.ToString();

    var logDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Logger");
    if (!Directory.Exists(logDirectory)) Directory.CreateDirectory(logDirectory);

    var filePath = Path.Combine(logDirectory, $"{now:yyyyMMdd}.txt");
    var logEntry = $"[{now:HH:mm:ss}] | {statusCode} | {ip} | {userName} | {method}: {path}{context.Request.QueryString}{Environment.NewLine}";

    try
    {
        await File.AppendAllTextAsync(filePath, logEntry);
    }
    catch { }
});

app.UseAuthentication();
app.UseAuthorization();
app.UseCookiePolicy();

app.MapControllerRoute(
    name: "areas",
    pattern: "{area:exists}/{controller=dashboard}/{action=index}/{id?}"
);

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=home}/{action=index}/{id?}"
);

//app.MapAreaControllerRoute(
//    name: "Admin",
//    areaName: "Admin",
//    pattern: "admin/{controller=dashboard}/{action=index}/{id?}"
//);

app.Run();
