namespace Web_QM.Models.ViewModels
{
    public class MaterialNode
    {
        public string Material { get; set; }
        public List<MachineCodeNode> MachineCodes { get; set; } = new List<MachineCodeNode>();
        public int MachineCount => MachineCodes.Count;
    }
}
