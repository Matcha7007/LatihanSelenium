using LatihanSelenium.Models;

namespace LatihanSelenium.Constants
{
	public class DriverConstant
	{
		public const string Chrome = "Chrome";
		public const string Edge = "Edge";
		public const string Firefox = "Firefox";
		public const string Safari = "Safari";
	}

	public class ModuleNameConstant
	{
		public const string PR = "Purchase Request";
		public const string RFQ = "Request For Quotation";
		public const string VQ = "Vendor Quotation";
		public const string VS = "Vendor Selection";
		public const string PO = "Purchase Order";
		public const string ROGS = "ROGS";
		public const string Invoice = "Invoice";
	}

	public class DataForConstant
	{
		public const string Submit = "Submit";
		public const string ReSubmit = "Re Submit";
		public const string Edit = "Edit";
		public const string SaveAsDraft = "Save As Draft";
		public const string SubmitDraft = "Submit Draft";
		public const string UpdateDraft = "Update Draft";
	}

	public class StatusConstant
	{
		public const string Success = "Success";
		public const string Fail = "Fail";
	}

	public class TestCaseConstant
	{
		public const string Submit = "Submit";
		public const string Approval = "Approval";
		public const string Edit = "Edit";
		public const string SearchList = "Search List";
		public const string SearchPendingTask = "Search Pending Task";
		public const string SaveAsDraft = "Save As Draft";
		public const string SubmitDraft = "Submit Draft";
		public const string UpdateDraft = "Update Draft";
		public const string DiscardDraft = "Discard Draft";
		public const string Close = "Close";
		public const string Cancel = "Cancel";
		
	}

	public class ApprovalConstant
	{
		public const string Approve = "Approve";
		public const string Reject = "Reject";
		public const string Return = "Return";		
		public const string Revise = "Revise";
		public const string ReSubmit = "Re Submit";
	}

	public class SheetConstant
	{
		public const string Config = "Config";
		public const string User = "User";
		public const string TestPlan = "TestPlan";
		public const string Task = "Task";
		public const string PR = "PurchaseRequest";
		public const string RFQ = "RequestForQuotation";
		public const string VQ = "VendorQuotation";
        public const string Invoice = "Invoice";
    }

	public class PrTypeConstant
	{
		public const string Catalog = "Catalog / Contract";
		public const string NonCatalog = "Non-Catalog / Non-Contract";
		public const string DirectPurchase = "Direct Purchase";
	}

    public class RfqTypeConstant
    {
        public const string DirectAppointment = "Direct Appointment";
        public const string Tender ="Tender";
    }

    public class InvoicePaymentTypeConstant
    {
        public const string Cash = "Cash";
        public const string BankTransfer = "Bank Transfer";
        public const string Other = "Other";
    }
    public class UrlConstant
	{
		private static string BaseUrl = $"{GlobalConfig.Config.Url}/app";
		private const string Create = "create";
		private const string List = "list?tabActive=list";
		private const string PendingTask = "list?tabActive=pending-task";

		#region Url
		public static string CreatePR = $"{BaseUrl}/purchase-request/{Create}";
		public static string ListPR = $"{BaseUrl}/purchase-request/{List}";
		public static string PendingTaskPR = $"{BaseUrl}/purchase-request/{PendingTask}";

        public static string CreateRFQ = $"{BaseUrl}/rfq/{Create}";
        public static string ListRFQ = $"{BaseUrl}/rfq/{List}";
        public static string PendingTaskRFQ = $"{BaseUrl}/rfq/{PendingTask}";

        public static string CreateInvoice = $"{BaseUrl}/invoice-receive/{Create}";
        public static string ListInvoice = $"{BaseUrl}/invoice-receive/{List}";
        public static string PendingTaskInvoice = $"{BaseUrl}/invoice-receive/{PendingTask}";

        public static string CreateVQ = $"{BaseUrl}/vendor-quotation/{Create}";
		public static string ListVQ = $"{BaseUrl}/vendor-quotation/{List}";
        #endregion
    }
}
