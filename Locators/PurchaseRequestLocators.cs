using OpenQA.Selenium;

namespace LatihanSelenium.Locators
{
	public class PurchaseRequestLocators
	{
		public static By SearchList = By.XPath("//div[@id='rc-tabs-0-panel-list']//input[@placeholder='Search']");

		public static By FieldPRList = By.XPath("//input[@id='form-create-pr_prTypeId']");
		public static By PRList = By.XPath("//div[@class='ant-select-item-option-content']");
		public static string UrlCreate = "https://sit2.quickacq.com/app/purchase-request/create";

		public static By AttachmentSection = By.XPath("//form//div[@class='ant-collapse-item']");
		public static By BtnAdd = By.XPath("//div[@class='purchase-request-section-attachments']//div[@class='wrap-input-upload-file']");
	}
}
