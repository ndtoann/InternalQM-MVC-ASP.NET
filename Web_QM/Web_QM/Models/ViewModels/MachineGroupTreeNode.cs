namespace Web_QM.Models.ViewModels
{
    public class MachineGroupTreeNode
    {
        public long MachineGroupId { get; set; }
        public string GroupName { get; set; }
        public string MachineType { get; set; }

        public List<MaterialNode> Materials { get; set; } = new List<MaterialNode>();

        public int MaterialCount => Materials.Count;
    }
}
