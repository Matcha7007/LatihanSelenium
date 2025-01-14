﻿using LatihanSelenium.Models;

namespace LatihanSelenium.Constants
{
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
	}

	public class SheetConstant
	{
		public const string Config = "Config";
		public const string User = "User";
		public const string TestPlan = "TestPlan";
		public const string Task = "Task";
		public const string PR = "PurchaseRequest";
	}

	public class PrTypeConstant
	{
		public const string Catalog = "Catalog / Contract";
		public const string NonCatalog = "Non-Catalog / Non-Contract";
		public const string DirectPurchase = "Direct Purchase";
	}

	public class UrlConstant
	{
		public static string CreatePR = $"{GlobalConfig.Config.Url}/app/purchase-request/list";
	}
}
