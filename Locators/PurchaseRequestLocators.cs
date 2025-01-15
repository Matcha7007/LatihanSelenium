using OpenQA.Selenium;

namespace LatihanSelenium.Locators
{
	public class PurchaseRequestLocators
	{
		public static By SearchList = By.XPath("//div[@id='rc-tabs-0-panel-list']//input[@placeholder='Search']");
		public static By RowListPrNumber = By.XPath("//tbody/tr[@data-row-key]/td[1]//div[@class='ellipsis']");

		public static By FieldPRList = By.XPath("//input[@id='form-create-pr_prTypeId']");
		public static By PRList = By.XPath("//div[@class='ant-select-item-option-content']");

		public static By AttachmentSection = By.XPath("//form//div[@class='ant-collapse-item']");
		public static By BtnAdd = By.XPath("//div[@class='purchase-request-section-attachments']//div[@class='wrap-input-upload-file']");
	}

	public class PurchaseRequestTaskLocators
	{
		public static By Search = By.XPath("");
		public static By RowTask = By.XPath("");
		public static By ApprovalRemark = By.XPath("");
		public static By BtnApprove = By.XPath("");
		public static By BtnRevise = By.XPath("");
		public static By BtnReject = By.XPath("");
		public static By BtnRsubmit = By.XPath("");
	}
}
