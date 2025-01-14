using LatihanSelenium.Constants;
using LatihanSelenium.Models;
using OpenQA.Selenium.Chrome;
using LatihanSelenium.Pages;

namespace LatihanSelenium.Utilities
{
	public static class UnitTestHandler
	{
		public static void RunUnitTest()
		{
			try
			{
				ResourceModels resourceData = ExcelHelper.ReadExcel(GlobalConfig.Config);

				foreach (var plan in resourceData.Testplans)
				{
					if (!string.IsNullOrEmpty(plan.Status))
						continue;
					Console.WriteLine("------------------------------------------------------");
					Console.WriteLine($"Running Unit Test {plan.TestCase} - {plan.ModuleName}");
					Console.WriteLine("------------------------------------------------------");
					ChromeDriver driver = new();

					try
					{
						SwitchCaseTest(GlobalConfig.Config, driver, plan, resourceData.Users, resourceData);
						Console.WriteLine("------------------------------------------------------");
						Console.WriteLine($"End of Unit Test {plan.TestCase} - {plan.ModuleName}.");
						Console.WriteLine("------------------------------------------------------");
					}
					catch (Exception ex)
					{
						Console.WriteLine("------------------------------------------------------");
						Console.WriteLine($"End of Unit Test {plan.TestCase} - {plan.ModuleName} is Fail. Msg : {ex.Message}");
						Console.WriteLine("------------------------------------------------------");
					}

					driver.Dispose();
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine("------------------------------------------------------");
				Console.WriteLine($"Run Unit Test Failed : {ex.Message}");
				Console.WriteLine("------------------------------------------------------");
			}
		}

		private static void SwitchCaseTest(AppConfig cfg, ChromeDriver driver, TestplanModels plan, List<UserModels> users, ResourceModels param)
		{
			try
			{
				UserModels user = users.Where(x => x.Username.Equals(plan.UserLogin)).FirstOrDefault()!;
				if (!plan.TestCase.Equals(TestCaseConstant.Approval))
				{
					driver.Manage().Window.FullScreen();
					LoginPages.Login(driver, user, cfg);
				}

				Console.WriteLine($"PLAN ID : {plan.TestCaseId} - MODULE NAME : {plan.ModuleName}");
				switch (plan.TestCase)
				{
					case TestCaseConstant.Submit:
					case TestCaseConstant.SubmitDraft:
						switch (plan.ModuleName)
						{
							case ModuleNameConstant.PR:
								PurchaseRequestPages.HandleTestCase(driver, cfg, plan, param);
								break;
							case ModuleNameConstant.RFQ:
								break;
							case ModuleNameConstant.VQ:
								break;
							case ModuleNameConstant.VS:
								break;
							case ModuleNameConstant.PO:
								break;
							case ModuleNameConstant.ROGS:
								break;
							case ModuleNameConstant.Invoice:
								break;
							default:
								break;
						}
						break;
					case TestCaseConstant.Approval:
						List<TaskModels> newTask = param.Tasks.Where(x => x.TestCaseId.Equals(plan.TestCaseId)).OrderBy(x => x.Sequence).ToList()!;
						//if (newTask.Count > 0)
							//tasks.HandleTaskAction(driver, cfg, plan, newTask, users, param);
						break;
					default: break;
				}

			}
			catch (Exception ex) { throw new Exception(ex.Message); }
			finally { driver.Dispose(); }
		}
	}
}
