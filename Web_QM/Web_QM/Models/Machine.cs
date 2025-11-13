using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class Machine
    {
        public long Id { get; set; }

        [Required(ErrorMessage = "Mã máy không được để trống")]
        public string MachineCode { get; set; }

        [Required(ErrorMessage = "Tên máy không được để trống")]
        public string MachineName { get; set; }

        public string? Department {  get; set; }

        public string? Version { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int? NumberOfAxes { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public int? NumberOfKnifeSocket {  get; set; }

        public string? SpindleSpeed { get; set; }

        public string? BottleTaper { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? MachineTableSizeX { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? MachineTableSizeY { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? MachineJourneyX { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? MachineJourneyY { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? MachineJourneyZ { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? WideSize { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? DeepSize { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? HighSize { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? OuterPairX { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? OuterPairZ { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? PairInsideX { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? PairInsideZ { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? TailstockX { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? TailstockZ { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? Weight { get; set; }

        public string? Picture { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? Price { get; set; }

        [Range(0, (double)decimal.MaxValue, ErrorMessage = "Giá trị phải lớn hơn hoặc bằng 0")]
        public decimal? MachineCapacity { get; set; }

        public string? MachineOrigin { get; set; }

        public string? Place {  get; set; }

        [Required(ErrorMessage = "Vui lòng chọn trạng thái")]
        public string Status { get; set; }

        public string? TypeMachine { get; set; }

        public string? Note { get; set; }

        public string? CreatedDate { get; set; }

        public string? UpdatedDate { get; set; }
    }
}
