using System.ComponentModel.DataAnnotations;

namespace Web_QM.Models
{
    public class ProductionProcessess
    {
        public long Id { get; set; }

        [Required]
        public string PartName { get; set; }

        public string? Picture { get; set; }

        [Required]
        public string WorkpieceSize { get; set; }

        [Required]
        public string Material { get; set; }

        public string CreatedBy { get; set; }

        public DateOnly CreatedDate { get; set; }

        [Range(0, int.MaxValue)]
        public int Version { get; set; }

        public string? UpdatedBy { get; set; }

        public DateOnly? UpdatedDate { get; set; }

        public string? Note { get; set; }

        public string? ModifiedContent { get; set; }
    }
}
