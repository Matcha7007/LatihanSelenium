using OpenQA.Selenium;

namespace LatihanSelenium.Utilities
{
	public static class PageHelpers
	{
		public static string SearchDocumentNumber(IWebDriver driver, string moduleName, By locatorSearch, By locatorRow, string param)
		{
			try
			{
				AutomationHelpers.ElementExist(driver, locatorSearch);

				AutomationHelpers.FillElementNonMandatory(driver, locatorSearch, param);
				AutomationHelpers.ElementExist(driver, locatorRow);
				return driver.FindElement(locatorRow).Text;
			}
			catch (Exception ex) 
			{ 
				throw new Exception($"Error trying to get the document number of {moduleName} with parameter {param}. Msg {ex}"); 
			}
		}
	}
}
