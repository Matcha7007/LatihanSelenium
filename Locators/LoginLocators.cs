using OpenQA.Selenium;

namespace LatihanSelenium.Locators
{
	public class LoginLocators
	{
		//public static string Username = "login_username";
		public static By Username = By.Id("login_username");
		public static By Password = By.Id("login_password");
		public static By BtnLogin = By.XPath("//button[@type='submit']");
	}
}
