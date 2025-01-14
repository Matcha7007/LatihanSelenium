namespace LatihanSelenium.Models
{
	public class PurchaseRequestListModels
	{
		public string SearchList { get; set; } = string.Empty;
	}

	public class PurchaseRequestModels : ModuleBase
	{
		public string PRSubject { get; set; } = string.Empty;
		public string CostCenter { get; set; } = string.Empty;
		public string PRType { get; set; } = string.Empty;
		public string RequiredDate { get; set; } = string.Empty;
		public string ProjectCode { get; set; } = string.Empty;
		public string Remarks { get; set; } = string.Empty;
		public virtual List<PRItem> GoodsItems { get; set; } = [];
		public virtual List<PRItem> ServicesItems { get; set; } = [];
		public int GoodsVATRate { get; set; }
		public int ServicesVATRate { get; set; }
		public List<string> AttachmentPaths { get; set; } = [];
	}

	public class PRItem
	{
		public string Name { get; set; } = string.Empty;
		public int Qty { get; set; }
		public int Price { get; set; }
}
