using OpenQA.Selenium;

namespace LatihanSelenium.Locators
{
	public class GlobalLocators
	{
		public static By AlertSuccess = By.XPath("//div[@role='status'][contains(text(), 'successfully.')]");
	}
}
