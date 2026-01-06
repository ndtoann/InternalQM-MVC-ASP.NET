namespace Web_QM.Models
{
    public class Opinion
    {
        public long Id { get; set; }

        public string Title { get; set; }

        public string Type { get; set; }

        public string Content { get; set; }

        public int? Status { get; set; } = 0;

        public long? CreatedBy { get; set; }

        public DateOnly? CreatedDate { get; set; }
    }
}
