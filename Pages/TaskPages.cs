using LatihanSelenium.Constants;
using LatihanSelenium.Locators;
using LatihanSelenium.Models;
using LatihanSelenium.Utilities;
using OpenQA.Selenium;

namespace LatihanSelenium.Pages
{
	public static class TaskPages
	{
		private static Dictionary<int, string> _dictionaryResult = [];

		public static IWebDriver HandleTaskAction(IWebDriver driver, AppConfig cfg, TestplanModels plan, List<TaskModels> param, List<UserModels> users, ResourceModels allParam)
		{
			_dictionaryResult.Clear();
			try
			{
				TestplanModels newPlan = plan;
				_dictionaryResult = [];
				foreach (TaskModels task in param)
				{
					driver.Dispose();
					driver = DriverManagerHelpers.NewWebDriver(cfg.Driver);
					UserModels user = users.Where(x => x.Username == task.Actor).FirstOrDefault()!;
					LoginPages.Login(driver, user, cfg);

					var dataTask = allParam.Testplans.Where(x => x.TestCaseId.Equals(plan.FromTestCaseId)).FirstOrDefault()!;
					task.DataNumber = dataTask.RequestNumber;
					plan.TestData = PlanHelper.CreateAddress(SheetConstant.Task, task.Row);

					Console.WriteLine($"TASK WITH TEST PLAN {task.TestCaseId} - {task.Sequence} - {task.Action} - {task.Actor} - {task.DataNumber}");
					string result = TaskHandle(driver, cfg, newPlan, task, allParam);

					if (result.Contains("Fail"))
					{
						_dictionaryResult.Add(task.Sequence, result);
						break;
					}
				}

				plan.Status = StatusConstant.Success;
				plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Successfully.";

				foreach (var item in _dictionaryResult.Values)
				{
					if (item.Contains("Fail") || item.Contains("Fail."))
					{
						plan.Status = StatusConstant.Fail;
						plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Fail. Error : One or more sequences have failed.";
						break;
					}
				}
			}
			catch (Exception ex)
			{
				plan.Status = StatusConstant.Fail;
				plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Fail. Error : {ex.Message}";
			}
			finally
			{
				ExcelHelpers.WriteAutomationResult(cfg, plan);
				Thread.Sleep(GlobalConfig.Config.WaitWriteResultInSecond);
			}
			return driver;
		}

		private static string TaskHandle(IWebDriver driver, AppConfig cfg, TestplanModels plan, TaskModels param, ResourceModels allParam)
		{
			string result = StatusConstant.Success;
			TaskLocatorModels locator = MappingTaskXPath(plan.ModuleName);
			try
			{
				bool isTaskExist = SearchTask(driver, param, plan.ModuleName, locator);

				if (isTaskExist)
				{
					if (param.Action.Equals(ApprovalConstant.ReSubmit, StringComparison.OrdinalIgnoreCase))
					{
						switch (plan.ModuleName)
						{
							case ModuleNameConstant.PR:
								PurchaseRequestModels prParam = allParam.PurchaseRequests.Where(x => x.TestCaseId.Equals(param.TestCaseId) && x.DataFor.Equals(DataForConstant.ReSubmit) && x.Sequence.Equals(param.Sequence)).FirstOrDefault()!;
								if (prParam != null)
									PurchaseRequestPages.HandleForm(driver, cfg, plan, prParam, false);
								break;							
							default:
								break;
						}
					}

					AutomationHelpers.ScrollToElement(driver, locator.ApprovalRemark);

					if (string.IsNullOrEmpty(param.Notes))
					{
						result = $"Notes parameter cannot be empty";
						param.Result = result;
						throw new Exception(result);
					}
					else
					{
						AutomationHelpers.FillElement(driver, locator.ApprovalRemark, param.Notes);

						if (param.Action.Equals(ApprovalConstant.Approve, StringComparison.OrdinalIgnoreCase))
							AutomationHelpers.ClickElement(driver, locator.BtnApprove);
						if (param.Action.Equals(ApprovalConstant.ReSubmit, StringComparison.OrdinalIgnoreCase))
							AutomationHelpers.ClickElement(driver, locator.BtnResubmit);
						if (param.Action.Equals(ApprovalConstant.Revise, StringComparison.OrdinalIgnoreCase))
							AutomationHelpers.ClickElement(driver, locator.BtnRevise);
						if (param.Action.Equals(ApprovalConstant.Reject, StringComparison.OrdinalIgnoreCase))
							AutomationHelpers.ClickElement(driver, locator.BtnReject);

						Thread.Sleep(500);
						string msg = AutomationHelpers.GetAlertMessage(driver, GlobalLocators.AlertSuccess);
						if (AutomationHelpers.ValidateAlert(msg, GlobalConfig.Config.SuccessMessages))
						{
							param.Result = StatusConstant.Success;
						}
						else
						{
							result = $"{msg}";
							param.Result = result;
							throw new Exception(result);
						}
					}
				}
				else
				{
					result = $"Something went wrong";
					param.Result = result;
					throw new Exception(result);
				}
			}
			catch (Exception ex)
			{
				result = $"Task {param.Action} {plan.ModuleName} is Fail. Error : {ex.Message}";
				param.Result = result;
			}
			finally
			{
				ExcelHelpers.WriteApproval<TaskModels>(cfg, param, plan.ModuleName);
				Thread.Sleep(GlobalConfig.Config.WaitWriteResultInSecond);
				driver.Dispose();
			}
			return result;
		}

		private static bool SearchTask(IWebDriver driver, TaskModels param, string moduleName, TaskLocatorModels locator)
		{
			int retryCount = 1;
			try
			{
				driver.Navigate().GoToUrl(locator.UrlPendingTask);

				AutomationHelpers.ElementExist(driver, locator.Search);

				while (true)
				{
					Console.WriteLine($"Try to Search {param.DataNumber} - {retryCount}");
					AutomationHelpers.FillElement(driver, locator.Search, param.DataNumber);

					try
					{
						Thread.Sleep(2000);
						AutomationHelpers.ClickElement(driver, locator.RowTask);
						return true;
					}
					catch
					{
						if (retryCount >= GlobalConfig.Config.MaxRetrySearchTask)
						{
							throw new Exception($"No data record with {locator.Module} No {param.DataNumber} after {GlobalConfig.Config.MaxRetrySearchTask} attempts.");
						}
						Thread.Sleep(GlobalConfig.Config.WaitRetrySearchTaskInMiliSecond);
						driver.FindElement(locator.Search).Clear();
						retryCount++;
					}
				}
			}
			catch (Exception ex)
			{
				throw new Exception($"Search Task {moduleName} failed: {ex.Message}");
			}
		}

		private static TaskLocatorModels MappingTaskXPath(string moduleName)
		{
			TaskLocatorModels xpath = new();
			switch (moduleName)
			{
				case ModuleNameConstant.PR:
					xpath = new()
					{
						Search = PurchaseRequestTaskLocators.Search,
						RowTask = PurchaseRequestTaskLocators.RowTask,
						ApprovalRemark = PurchaseRequestTaskLocators.ApprovalRemark,
						BtnApprove = PurchaseRequestTaskLocators.BtnApprove,
						BtnRevise = PurchaseRequestTaskLocators.BtnRevise,
						BtnReject = PurchaseRequestTaskLocators.BtnReject,
						UrlPendingTask = UrlConstant.PendingTaskPR,
						Module = ModuleNameConstant.PR
					};
					break;
				case ModuleNameConstant.RFQ:
					xpath = new()
					{
						//Search = PurchaseRequestTaskLocators.Search,
						//RowTask = PurchaseRequestTaskLocators.RowTask,
						//ApprovalRemark = PurchaseRequestTaskLocators.ApprovalRemark,
						//BtnApprove = PurchaseRequestTaskLocators.BtnApprove,
						//BtnRevise = PurchaseRequestTaskLocators.BtnRevise,
						//BtnReject = PurchaseRequestTaskLocators.BtnReject,
						//UrlPendingTask = UrlConstant.PendingTaskPR,
						Module = ModuleNameConstant.RFQ
					};
					break;
				case ModuleNameConstant.VS:
					xpath = new()
					{
						//Search = PurchaseRequestTaskLocators.Search,
						//RowTask = PurchaseRequestTaskLocators.RowTask,
						//ApprovalRemark = PurchaseRequestTaskLocators.ApprovalRemark,
						//BtnApprove = PurchaseRequestTaskLocators.BtnApprove,
						//BtnRevise = PurchaseRequestTaskLocators.BtnRevise,
						//BtnReject = PurchaseRequestTaskLocators.BtnReject,
						//UrlPendingTask = UrlConstant.PendingTaskPR,
						Module = ModuleNameConstant.VS
					};
					break;
				case ModuleNameConstant.PO:
					xpath = new()
					{
						//Search = PurchaseRequestTaskLocators.Search,
						//RowTask = PurchaseRequestTaskLocators.RowTask,
						//ApprovalRemark = PurchaseRequestTaskLocators.ApprovalRemark,
						//BtnApprove = PurchaseRequestTaskLocators.BtnApprove,
						//BtnRevise = PurchaseRequestTaskLocators.BtnRevise,
						//BtnReject = PurchaseRequestTaskLocators.BtnReject,
						//UrlPendingTask = UrlConstant.PendingTaskPR,
						Module = ModuleNameConstant.PO
					};
					break;
				case ModuleNameConstant.Invoice:
					xpath = new()
					{
						//Search = PurchaseRequestTaskLocators.Search,
						//RowTask = PurchaseRequestTaskLocators.RowTask,
						//ApprovalRemark = PurchaseRequestTaskLocators.ApprovalRemark,
						//BtnApprove = PurchaseRequestTaskLocators.BtnApprove,
						//BtnRevise = PurchaseRequestTaskLocators.BtnRevise,
						//BtnReject = PurchaseRequestTaskLocators.BtnReject,
						//UrlPendingTask = UrlConstant.PendingTaskPR,
						Module = ModuleNameConstant.Invoice
					};
					break;
				default: break;
			}
			return xpath;
		}
	}
}
