using LatihanSelenium.Constants;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Safari;

namespace LatihanSelenium.Utilities
{
	public static class DriverManagerHelpers
	{
		public static IWebDriver NewWebDriver(string driverType)
		{
			switch (driverType)
			{
				case DriverConstant.Edge:
					return new EdgeDriver();
				case DriverConstant.Firefox:
					return new FirefoxDriver();
				case DriverConstant.Safari:
					return new SafariDriver();
				default:
					return new ChromeDriver();
			}
		}
	}
}
