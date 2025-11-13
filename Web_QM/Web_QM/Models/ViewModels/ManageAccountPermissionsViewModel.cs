namespace Web_QM.Models.ViewModels
{
    public class ManageAccountPermissionsViewModel
    {
        public List<Account> AvailableAccounts { get; set; }
        public Dictionary<string, List<Permission>> GroupedPermissions { get; set; }
        public long SelectedAccountId { get; set; }
    }
}
