namespace Web_QM.Models.ViewModels
{
    public class MachineToleranceDto
    {
        public string Type { get; set; }
        public decimal? LatestParameter { get; set; }
        public DateOnly? LatestDate { get; set; }
        public long? LatestId { get; set; }
    }
}
