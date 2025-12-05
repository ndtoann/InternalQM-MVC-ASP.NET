using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models.ViewModels
{
    public class ErrorDataView
    {
        public long Id { get; set; }

        public string OrderNo { get; set; }

        public string PartName { get; set; }

        public int QtyOrder { get; set; }

        public int QtyNG { get; set; }

        public DateOnly DateMonth { get; set; }

        public string ErrorDetected { get; set; }

        public string ErrorType { get; set; }

        public string NCC { get; set; }

        public string? EmployeeCode { get; set; }

        public string? Department { get; set; }

        public DateOnly? ErrorCompletionDate { get; set; }
    }
}
