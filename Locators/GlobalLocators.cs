using OpenQA.Selenium;

namespace LatihanSelenium.Locators
{
	public class GlobalLocators
	{
		private static string ConfirmAlertBase = "//div[@class='custom-alert-content-container']";
		public static By AlertSuccess = By.XPath("//div[@role='status'][contains(text(), 'successfully.')]");
		public static By ConfirmAlertBox = By.XPath($"{ConfirmAlertBase}");
		public static By ConfirmAlertYes = By.XPath($"{ConfirmAlertBase}//button[not(contains(@class, 'ant-btn-dangerous'))]");
	}
}
