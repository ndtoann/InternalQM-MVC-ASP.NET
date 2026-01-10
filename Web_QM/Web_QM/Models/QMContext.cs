using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace Web_QM.Models;

public partial class QMContext : DbContext
{
    public QMContext()
    {
    }

    public QMContext(DbContextOptions<QMContext> options)
        : base(options)
    {
    }

    public virtual DbSet<Account> Accounts { get; set; }

    public virtual DbSet<Permission> Permissions { get; set; }
    public virtual DbSet<AccountPermission> AccountPermissions { get; set; }

    public virtual DbSet<Employee> Employees { get; set; }

    public virtual DbSet<Department> Departments { get; set; }

    public virtual DbSet<ExamPeriodicAnswer> ExamPeriodicAnswers { get; set; }

    public virtual DbSet<ExamTrainingAnswer> ExamTrainingAnswers { get; set; }

    public virtual DbSet<EmployeeTrainingResult> EmployeeTrainingResults { get; set; }

    public virtual DbSet<TestPractice> TestPractices { get; set; }

    public virtual DbSet<TestPracticeDetail> TestPracticeDetails { get; set; }

    public virtual DbSet<ExamPeriodic> ExamPeriodics { get; set; }

    public virtual DbSet<ExamTraining> ExamTrainings { get; set; }

    public virtual DbSet<Feedback> Feedbacks { get; set; }

    public virtual DbSet<Question> Questions { get; set; }

    public virtual DbSet<QuestionTraining> QuestionTrainings { get; set; }

    public virtual DbSet<Training> Trainings { get; set; }

    public virtual DbSet<ExamTrialRun> ExamTrialRuns { get; set; }

    public virtual DbSet<QuestionTrialRun> QuestionTrialRuns { get; set; }

    public virtual DbSet<ExamTrialRunAnswer> ExamTrialRunAnswers { get; set; }

    public virtual DbSet<Notification> Notifications { get; set; }

    public virtual DbSet<ErrorData> ErrorDatas { get; set; }

    public virtual DbSet<EmployeeWorkHistory> EmployeeWorkHistories { get; set; }

    public virtual DbSet<Productivity> Productivities { get; set; }
    public virtual DbSet<SawingPerformance> SawingPerformances { get; set; }

    public virtual DbSet<Kaizen> Kaizens { get; set; }

    public virtual DbSet<Violation5S> Violation5S { get; set; }

    public virtual DbSet<EmployeeViolation5S> EmployeeViolation5S { get; set; }
    public virtual DbSet<Machine> Machines { get; set; }
    public virtual DbSet<MachineParameter> MachineParameters { get; set; }
    public virtual DbSet<MachineGroup> MachineGroups { get; set; }
    public virtual DbSet<Machine_MG> Machine_MG { get; set; }
    public virtual DbSet<EquipmentRepairHistory> EquipmentRepairHistories { get; set; }
    public virtual DbSet<MachineMaintenance> MachineMaintenances { get; set; }
    public virtual DbSet<ReplacementEquipmentAndSupplies> ReplacementEquipmentAndSupplies { get; set; }

    public virtual DbSet<Salary> Salaries { get; set; }
    public virtual DbSet<MonthlyPayroll> MonthlyPayroll { get; set; }
    public virtual DbSet<ComplaintSalary> ComplaintSalary { get; set; }

    public virtual DbSet<Timesheet> Timesheets { get; set; }
    public virtual DbSet<Timekeeping> Timekeepings { get; set; }

    public virtual DbSet<Opinion> Opinions { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        => optionsBuilder.UseSqlServer("Server=192.168.10.251,1433;Database=QM_DaoTao;User Id=daotao01;Password=12341234;Trusted_Connection=False;MultipleActiveResultSets=true;Encrypt=False;");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<Account>(entity =>
        {
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.Password).HasMaxLength(500);
            entity.Property(e => e.Salt).HasMaxLength(500);
            entity.Property(e => e.StaffCode).HasMaxLength(20);
            entity.Property(e => e.StaffName).HasMaxLength(50);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
            entity.Property(e => e.UserName).HasMaxLength(50);
        });

        modelBuilder.Entity<Permission>(entity =>
        {
            entity.Property(e => e.Module).HasMaxLength(50);
            entity.Property(e => e.ClaimValue).HasMaxLength(100);
            entity.Property(e => e.Description).HasMaxLength(100);
        });

        modelBuilder.Entity<AccountPermission>()
            .HasKey(rp => new { rp.AccountId, rp.PermissionId });

        modelBuilder.Entity<Employee>(entity =>
        {
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.Department).HasMaxLength(50);
            entity.Property(e => e.EmployeeCode).HasMaxLength(50);
            entity.Property(e => e.EmployeeName).HasMaxLength(50);
            entity.Property(e => e.Gender).HasMaxLength(5);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
            entity.Property(e => e.Position).HasMaxLength(50);
            entity.Property(e => e.Avatar).HasMaxLength(255);
        });

        modelBuilder.Entity<ExamPeriodicAnswer>(entity =>
        {
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.EmployeeCode).HasMaxLength(10);
            entity.Property(e => e.EmployeeName).HasMaxLength(50);
            entity.Property(e => e.ListAnswer).HasMaxLength(500);
            entity.Property(e => e.TlPoint).HasDefaultValue(0);
            entity.Property(e => e.TnPoint).HasDefaultValue(0);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
        });

        modelBuilder.Entity<ExamTrainingAnswer>(entity =>
        {
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.EmployeeCode).HasMaxLength(10);
            entity.Property(e => e.EmployeeName).HasMaxLength(50);
            entity.Property(e => e.ListAnswer).HasMaxLength(500);
            entity.Property(e => e.TlPoint).HasDefaultValue(0);
            entity.Property(e => e.TnPoint).HasDefaultValue(0);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
        });

        modelBuilder.Entity<EmployeeTrainingResult>(entity =>
        {
            entity.Property(e => e.EvaluationPeriod).HasMaxLength(100);
        });

        modelBuilder.Entity<ExamPeriodic>(entity =>
        {
            entity.Property(e => e.ExamName).HasMaxLength(500);
        });

        modelBuilder.Entity<ExamTraining>(entity =>
        {
            entity.Property(e => e.ExamName).HasMaxLength(500);
        });

        modelBuilder.Entity<Feedback>(entity =>
        {
            entity.Property(e => e.FeedbackerName).HasMaxLength(50);
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
        });

        modelBuilder.Entity<Question>(entity =>
        {
            entity.Property(e => e.CorrectOption).HasMaxLength(10);
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
        });

        modelBuilder.Entity<QuestionTraining>(entity =>
        {
            entity.Property(e => e.CorrectOption).HasMaxLength(10);
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
        });

        modelBuilder.Entity<Training>(entity =>
        {
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.TrainingName).HasMaxLength(100);
            entity.Property(e => e.Type).HasMaxLength(20);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
        });

        modelBuilder.Entity<TestPractice>(entity =>
        {
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.TestName).HasMaxLength(50);
            entity.Property(e => e.TestLevel).HasMaxLength(20);
            entity.Property(e => e.PartName).HasMaxLength(50);
            entity.Property(e => e.Result).HasMaxLength(10);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
        });

        modelBuilder.Entity<TestPracticeDetail>(entity =>
        {
            entity.Property(e => e.OperationName).HasMaxLength(10);
        });

        modelBuilder.Entity<ExamTrialRun>(entity =>
        {
            entity.Property(e => e.ExamName).HasMaxLength(500);
            entity.Property(e => e.TestLevel).HasMaxLength(20);
        });

        modelBuilder.Entity<QuestionTrialRun>(entity =>
        {
            entity.Property(e => e.CorrectOption).HasMaxLength(10);
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
        });

        modelBuilder.Entity<ExamTrialRunAnswer>(entity =>
        {
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.EmployeeCode).HasMaxLength(10);
            entity.Property(e => e.EmployeeName).HasMaxLength(50);
            entity.Property(e => e.ListAnswer).HasMaxLength(500);
            entity.Property(e => e.MultipleChoiceCorrect).HasDefaultValue(0);
            entity.Property(e => e.MultipleChoiceInCorrect).HasDefaultValue(0);
            entity.Property(e => e.MultipleChoiceFail).HasDefaultValue(0);
            entity.Property(e => e.EssayCorrect).HasDefaultValue(0);
            entity.Property(e => e.EssayInCorrect).HasDefaultValue(0);
            entity.Property(e => e.EssayFail).HasDefaultValue(0);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
        });

        modelBuilder.Entity<Notification>(entity =>
        {
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.Message).HasMaxLength(200);
        });

        modelBuilder.Entity<ErrorData>(entity =>
        {
            entity.Property(e => e.OrderNo).HasMaxLength(500);
            entity.Property(e => e.PartName).HasMaxLength(100);
            entity.Property(e => e.ErrorDetected).HasMaxLength(50);
            entity.Property(e => e.ErrorType).HasMaxLength(50);
            entity.Property(e => e.ErrorCause).HasMaxLength(50);
            entity.Property(e => e.ErrorContent).HasMaxLength(500);
            entity.Property(e => e.ToleranceAssessment).HasMaxLength(20);
            entity.Property(e => e.NCC).HasMaxLength(10);
            entity.Property(e => e.EmployeeCode).HasMaxLength(50);
            entity.Property(e => e.Department).HasMaxLength(20);
            entity.Property(e => e.RemedialMeasures).HasMaxLength(50);
            entity.Property(e => e.TimeWriteError).HasMaxLength(50);
            entity.Property(e => e.ReviewNnds).HasMaxLength(500);
        });

        modelBuilder.Entity<Department>(entity =>
        {
            entity.Property(e => e.DepartmentName).HasMaxLength(100);
            entity.Property(e => e.Note).HasMaxLength(500);
        });

        modelBuilder.Entity<Productivity>(entity =>
        {
            entity.Property(e => e.EmployeeCode).HasMaxLength(50);
            entity.Property(e => e.EmployeeName).HasMaxLength(50);
            entity.Property(e => e.ProductivityScore).HasColumnType("decimal(5, 2)");
        });

        modelBuilder.Entity<SawingPerformance>(entity =>
        {
            entity.Property(e => e.EmployeeCode).HasMaxLength(50);
            entity.Property(e => e.EmployeeName).HasMaxLength(50);
            entity.Property(e => e.SalesAmountUSD).HasColumnType("decimal(10, 2)");
            entity.Property(e => e.SalesRate).HasColumnType("decimal(10, 2)");
        });

        modelBuilder.Entity<Kaizen>(entity =>
        {
            entity.Property(e => e.EmployeeCode).HasMaxLength(50);
            entity.Property(e => e.EmployeeName).HasMaxLength(50);
            entity.Property(e => e.Department).HasMaxLength(50);
            entity.Property(e => e.AppliedDepartment).HasMaxLength(50);
            entity.Property(e => e.ImprovementGoal).HasMaxLength(50);
            entity.Property(e => e.TeamLeaderRating).HasMaxLength(10);
            entity.Property(e => e.ManagementReview).HasMaxLength(10);
            entity.Property(e => e.Picture).HasMaxLength(255);
            entity.Property(e => e.Deadline).HasMaxLength(100);
            entity.Property(e => e.StartTime).HasMaxLength(100);
            entity.Property(e => e.CurrentStatus).HasMaxLength(500);
        });

        modelBuilder.Entity<EmployeeWorkHistory>(entity =>
        {
            entity.Property(e => e.EmployeeCode).HasMaxLength(50);
            entity.Property(e => e.EmployeeName).HasMaxLength(50);
            entity.Property(e => e.Department).HasMaxLength(50);
        });

        modelBuilder.Entity<EmployeeViolation5S>(entity =>
        {
            entity.Property(e => e.EmployeeCode).HasMaxLength(50);
        });

        modelBuilder.Entity<Machine>(entity =>
        {
            entity.Property(e => e.MachineCode).HasMaxLength(50);
            entity.Property(e => e.MachineName).HasMaxLength(100);
            entity.Property(e => e.Department).HasMaxLength(50);
            entity.Property(e => e.Version).HasMaxLength(100);
            entity.Property(e => e.SpindleSpeed).HasMaxLength(50);
            entity.Property(e => e.BottleTaper).HasMaxLength(50);
            entity.Property(e => e.MachineOrigin).HasMaxLength(50);
            entity.Property(e => e.Place).HasMaxLength(20);
            entity.Property(e => e.Picture).HasMaxLength(200);
            entity.Property(e => e.Status).HasMaxLength(100);
            entity.Property(e => e.TypeMachine).HasMaxLength(100);
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
            entity.Property(e => e.MachineTableSizeX).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.MachineTableSizeY).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.MachineJourneyX).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.MachineJourneyY).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.MachineJourneyZ).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.WideSize).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.DeepSize).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.HighSize).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.OuterPairX).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.OuterPairZ).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.PairInsideX).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.PairInsideZ).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.TailstockX).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.TailstockZ).HasColumnType("decimal(18, 2)");
            entity.Property(e => e.Weight).HasColumnType("decimal(4, 2)");
            entity.Property(e => e.Price).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.MachineCapacity).HasColumnType("decimal(10, 3)");
        });

        modelBuilder.Entity<MachineParameter>(entity => {
            entity.Property(e => e.MachineCode).HasMaxLength(50);
            entity.Property(e => e.Type).HasMaxLength(100);
            entity.Property(e => e.Parameters).HasColumnType("decimal(6, 3)");
        });

        modelBuilder.Entity<MachineGroup>(entity => {
            entity.Property(e => e.GroupName).HasMaxLength(50);
            entity.Property(e => e.MachineType).HasMaxLength(200);
            entity.Property(e => e.Standard).HasMaxLength(500);
        });

        modelBuilder.Entity<Machine_MG>(entity => {
            entity.Property(e => e.MachineCode).HasMaxLength(50);
        });

        modelBuilder.Entity<Machine_MG>()
            .HasKey(mmg => new { mmg.MachineCode, mmg.MachineGroupId, mmg.Material });

        modelBuilder.Entity<EquipmentRepairHistory>(entity => {
            entity.Property(e => e.EquipmentCode).HasMaxLength(50);
            entity.Property(e => e.EquipmentName).HasMaxLength(200);
            entity.Property(e => e.RemedialStaff).HasMaxLength(500);
            entity.Property(e => e.RecipientOfRepairedDevice).HasMaxLength(100);
            entity.Property(e => e.RepairCosts).HasMaxLength(500);
            entity.Property(e => e.Department).HasMaxLength(50);
        });

        modelBuilder.Entity<ReplacementEquipmentAndSupplies>(entity => {
            entity.Property(e => e.EquipmentAndlSupplies).HasMaxLength(500);
            entity.Property(e => e.FilePdf).HasMaxLength(200);
        });

        modelBuilder.Entity<MachineMaintenance>(entity => {
            entity.Property(e => e.MachineCode).HasMaxLength(50);
            entity.Property(e => e.MaintenanceStaff).HasMaxLength(200);
        });

        modelBuilder.Entity<Salary>(entity =>
        {
            entity.Property(e => e.EmployeeCode).HasMaxLength(50);
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
            entity.Property(e => e.BaseSalary).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.InsuranceSalary).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.MealAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.DailyResponsibilityPay).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.FuelAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.ExternalWorkAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.HousingSubsidy).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.DiligencePay).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.NoViolationBonus).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.HazardousAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.CNCStressAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.SeniorityAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.CertificateAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.WorkingEnvironmentAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.JobPositionAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.MachineTestRunAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.RVFMachineMeasurementAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.StainlessSteelCleaningAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.FactoryGuardAllowance).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.HeavyDutyAllowance).HasColumnType("decimal(18, 0)");
        });

        modelBuilder.Entity<MonthlyPayroll>(entity =>
        {
            entity.Property(e => e.EmployeeCode).HasMaxLength(50);
            entity.Property(e => e.DateMonth).HasMaxLength(50);
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
            entity.Property(e => e.BusinessTripAndPhoneFee).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.Penalty5S).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.UnionFee).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.PIT).HasColumnType("decimal(18, 0)");
            entity.Property(e => e.TotalSalary).HasColumnType("decimal(18, 0)");
        });

        modelBuilder.Entity<ComplaintSalary>(entity =>
        {
            entity.Property(e => e.CreatedDate).HasMaxLength(50);
            entity.Property(e => e.UpdatedDate).HasMaxLength(50);
        });

        modelBuilder.Entity<Timekeeping>(entity => {
            entity.Property(e => e.Shift).HasMaxLength(20);
            entity.Property(e => e.TotalHours).HasColumnType("decimal(5, 2)");
        });

        modelBuilder.Entity<Opinion>(entity => {
            entity.Property(e => e.Title).HasMaxLength(255);
            entity.Property(e => e.Type).HasMaxLength(50);
            entity.Property(e => e.Img).HasMaxLength(255);
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
