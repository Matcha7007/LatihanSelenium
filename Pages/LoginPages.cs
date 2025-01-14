using LatihanSelenium.Locators;
using LatihanSelenium.Models;
using LatihanSelenium.Utilities;
using OpenQA.Selenium;

namespace LatihanSelenium.Pages
{
	public static class LoginPages
	{
		public static void Login(IWebDriver driver, UserModels param, AppConfig cfg)
		{
			try
			{
				//Untuk mengarahkan ke url di browser
				driver.Navigate().GoToUrl(cfg.Url);

				//Input username
				AutomationHelpers.FillElement(driver, LoginLocators.Username, param.Username);

				//Input password
				AutomationHelpers.FillElement(driver, LoginLocators.Password, param.Password);

				//Klik button login
				AutomationHelpers.ClickElement(driver, LoginLocators.BtnLogin);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
			}
		}
	}
}
