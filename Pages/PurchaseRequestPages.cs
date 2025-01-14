using LatihanSelenium.Constants;
using LatihanSelenium.Locators;
using LatihanSelenium.Models;
using LatihanSelenium.Utilities;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace LatihanSelenium.Pages
{
	public static class PurchaseRequestPages
	{
		public static void HandleTestCase(ChromeDriver driver, AppConfig cfg, TestplanModels plan, ResourceModels param)
		{
			try
			{
				switch (plan.TestCase)
				{
					case TestCaseConstant.Submit:
						PurchaseRequestModels dataSubmit = param.PurchaseRequests.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataForConstant.Submit)).FirstOrDefault()!;
						plan.TestData = PlanHelper.CreateAddress(SheetConstant.PR, dataSubmit.Row);
						FormPR(driver, cfg, plan, dataSubmit, true);
						break;
					case TestCaseConstant.SaveAsDraft:
						PurchaseRequestModels dataDraft = param.PurchaseRequests.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataForConstant.SaveAsDraft)).FirstOrDefault()!;
						plan.TestData = PlanHelper.CreateAddress(SheetConstant.PR, dataDraft.Row);
						FormPR(driver, cfg, plan, dataDraft);
						break;
					default: break;
				}
			}
			finally
			{
				driver.Dispose();
			}
		}

		public static void SearchListPR(IWebDriver driver)
		{
			try
			{
				PurchaseRequestListModels param = new()
				{
					SearchList = "PR/2025/01/0002"
				};

				AutomationHelpers.ElementExist(driver, DashboardLocators.Topbar);
				//Mengarahkan ke halaman PR
				driver.Navigate().GoToUrl(UrlConstant.CreatePR);

				//Input Search List
				AutomationHelpers.FillElement(driver, PurchaseRequestLocators.SearchList, param.SearchList);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
			}
		}

		public static void FormPR(ChromeDriver driver, AppConfig cfg, TestplanModels plan, PurchaseRequestModels param, bool isSubmit = false)
		{
			try
			{
				AutomationHelpers.ElementExist(driver, DashboardLocators.Topbar);

				//Mengarahkan ke halaman Create PR
				driver.Navigate().GoToUrl(UrlConstant.CreatePR);

				//Select PR Type
				AutomationHelpers.SelectElement(driver, PurchaseRequestLocators.FieldPRList, PurchaseRequestLocators.PRList, param.PRType);

				//buka item section

				//Add Goods Items
				foreach (PRItem goodsItem in param.GoodsItems)
				{
					//klik btn add goods item
					//search item by goodsItem.Name
					//klik add
					//klik close
					//isi qty

					//kondisi jika pr type non catalog
					if (param.PRType.Equals(PrTypeConstant.NonCatalog))
					{
						//bisa edit price
					}
				}

				//Upload Attachment
				foreach (string path in param.AttachmentPaths)
				{
					AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.AttachmentSection);
					AutomationHelpers.UploadFile(driver, PurchaseRequestLocators.BtnAdd, path);
				}

				if (isSubmit)
				{
					//klik button submit
				}
				else
				{
					//klik button save as draft
				}
			}
			catch (Exception ex) 
			{ 
				Console.WriteLine(ex);
			}
			
		}
	}
}
