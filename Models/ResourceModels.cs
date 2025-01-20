namespace LatihanSelenium.Models
{
	public class ResourceModels
	{
		public virtual List<UserModels> Users { get; set; } = [];
		public virtual List<TestplanModels> Testplans { get; set; } = [];
		public virtual List<TaskModels> Tasks { get; set; } = [];
		public virtual List<PurchaseRequestModels> PurchaseRequests { get; set; } = [];
        public virtual List<RequestForQuotationModels> RequestForQuotations { get; set; } = [];
    }
}
