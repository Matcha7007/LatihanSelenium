using InputSimulatorEx;
using InputSimulatorEx.Native;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

using System.Runtime.InteropServices;

namespace LatihanSelenium.Utilities
{
	public class AutomationHelpers
	{
		// Import necessary Windows API functions
		[DllImport("user32.dll", SetLastError = true)]
		private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

		[DllImport("user32.dll", SetLastError = true)]
		private static extern bool SetForegroundWindow(IntPtr hWnd);

		public static void FillElement(IWebDriver driver, By locator, string param)
		{
			try
			{
				ElementExist(driver, locator);
				driver.FindElement(locator).SendKeys(param);
			}
			catch (Exception ex) 
			{
				throw new Exception(ex.ToString());
			}
		}

		public static void FillElementNonMandatory(IWebDriver driver, By locator, string param)
		{
			if (!string.IsNullOrEmpty(param)) FillElement(driver, locator, param);
		}

		public static void SelectElement(IWebDriver driver, By locatorField, By locatorList, string param)
		{
			try
			{
				ElementExist(driver, locatorField);
				driver.FindElement(locatorField).Click();
				ElementExist(driver, locatorList);
				var options = driver.FindElements(locatorList);

				foreach (var option in options)
				{
					if (option.Text.Contains(param, StringComparison.OrdinalIgnoreCase))
					{
						option.Click();
					}
				}
			}
			catch (Exception ex)
			{
				throw new Exception(ex.ToString());
			}
		}

		public static void SelectElementNonMandatory(IWebDriver driver, By locatorField, By locatorList, string param)
		{
			if (!string.IsNullOrEmpty(param)) SelectElement(driver, locatorField, locatorList, param);
		}

		public static void ClickElement(IWebDriver driver, By locator)
		{
			try
			{
				WebDriverWait wait = new(driver, TimeSpan.FromSeconds(10));
				wait.Until(ExpectedConditions.ElementToBeClickable(locator));

				driver.FindElement(locator).Click();
			}
			catch (WebDriverTimeoutException ex)
			{
				throw new Exception(ex.ToString());
			}
			catch (ElementNotInteractableException ex)
			{
				throw new Exception(ex.ToString());
			}
			catch (Exception ex)
			{
				throw new Exception(ex.ToString());
			}
		}

		public static void ElementExist(IWebDriver driver, By locator)
		{
			try
			{
				WebDriverWait wait = new(driver, TimeSpan.FromSeconds(10));
				wait.Until(ExpectedConditions.ElementExists(locator));
			}
			catch (WebDriverTimeoutException ex)
			{
				throw new Exception(ex.ToString());
			}
			catch (NoSuchElementException ex)
			{
				throw new Exception(ex.ToString());
			}
			catch (Exception ex)
			{
				throw new Exception(ex.ToString());
			}
		}

		public static void UploadFile(IWebDriver driver, By locator, string filePath, string title = "Open")
		{
			try
			{
				if (string.IsNullOrEmpty(filePath))
				{
					throw new Exception($"File path parameter is mandatory for the element {locator}.");
				}
				Thread.Sleep(1000);
				var btnUpload = driver.FindElement(locator);

				OpenQA.Selenium.Interactions.Actions action = new(driver);
				action.Click(btnUpload).Build().Perform();

				Thread.Sleep(TimeSpan.FromSeconds(3));

				var dialogHWnd = FindWindow(null!, title); // Here goes the title of the dialog window
				var setFocus = SetForegroundWindow(dialogHWnd);
				if (setFocus)
				{
					var sim = new InputSimulator();
					sim.Keyboard
						.KeyPress(VirtualKeyCode.RETURN)
						.Sleep(150)
						.TextEntry(filePath)
						.Sleep(1000)
						.KeyPress(VirtualKeyCode.RETURN);
				}
				Thread.Sleep(1000);
			}
			catch (Exception ex)
			{
				throw new Exception($"Fail Upload File : {ex.Message}");
			}
		}

		public static void UploadFileNonMandatory(IWebDriver driver, By locator, string filePath, string title = "Open")
		{
			if (!string.IsNullOrEmpty(filePath)) UploadFile(driver, locator, filePath, title);
		}

		public static void ClickInteractableElementUsingJS(EdgeDriver driver, By locator)
		{
			try
			{
				Thread.Sleep(500);
				IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
				js.ExecuteScript("arguments[0].click();", locator);
				Console.WriteLine($"Click button {locator?.ToString()}");
			}
			catch (Exception ex) { Console.WriteLine($"Error Click button {locator?.ToString()}"); throw new Exception(ex.Message); }
		}
	}
}
