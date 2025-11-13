namespace Web_QM.Models.ViewModels
{
    public class PermissionUpdateModel
    {
        public long SelectedAccountId { get; set; }
        public List<long> AssignedPermissionIds { get; set; }
    }
}
