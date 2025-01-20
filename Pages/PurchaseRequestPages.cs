using LatihanSelenium.Constants;
using LatihanSelenium.Locators;
using LatihanSelenium.Models;
using LatihanSelenium.Utilities;
using OpenQA.Selenium;

namespace LatihanSelenium.Pages
{
	public static class PurchaseRequestPages
	{
		public static void HandleTestCase(IWebDriver driver, AppConfig cfg, TestplanModels plan, ResourceModels param)
		{
			try
			{
				switch (plan.TestCase)
				{
					case TestCaseConstant.Submit:
						PurchaseRequestModels dataSubmit = param.PurchaseRequests.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataForConstant.Submit)).FirstOrDefault()!;
						plan.TestData = PlanHelper.CreateAddress(SheetConstant.PR, dataSubmit.Row);
						HandleForm(driver, cfg, plan, dataSubmit, true);
						break;
					case TestCaseConstant.SaveAsDraft:
						PurchaseRequestModels dataDraft = param.PurchaseRequests.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataForConstant.SaveAsDraft)).FirstOrDefault()!;
						plan.TestData = PlanHelper.CreateAddress(SheetConstant.PR, dataDraft.Row);
						HandleForm(driver, cfg, plan, dataDraft);
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

		public static void HandleForm(IWebDriver driver, AppConfig cfg, TestplanModels plan, PurchaseRequestModels param, bool isSubmit = false, bool isResubmit = false)
		{
			try
			{
				if (!isResubmit)
				{
					AutomationHelpers.NavigateTo(driver, UrlConstant.CreatePR);
				}

                //Create PR Header Section
                //Fill in the PR Subject
                AutomationHelpers.FillElement(driver, PurchaseRequestLocators.FillSubjectPR, param.PRSubject);

                //Select Cost Center 
                AutomationHelpers.SelectElement(driver, PurchaseRequestLocators.CostCenterPRField, PurchaseRequestLocators.CostCenterPRList, param.CostCenter);

                //Select PR Type
                AutomationHelpers.SelectElement(driver, PurchaseRequestLocators.TypePRField, PurchaseRequestLocators.TypePRList, param.PRType);

                //Fill in RequiredDate
                AutomationHelpers.FillElementNonMandatory(driver, PurchaseRequestLocators.RequiredDatePR, param.RequiredDate);

                //Fill in ProjectCode
                AutomationHelpers.FillElementNonMandatory(driver, PurchaseRequestLocators.ProjectCodePR, param.ProjectCode);

                //Fill in Remarks
                AutomationHelpers.FillElementNonMandatory(driver, PurchaseRequestLocators.RemarksPR, param.Remarks);

                //Create PR Items Section
                //buka item section
                AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.ItemBtn);

                //Add Goods Items
                foreach (PRItem goodsItem in param.GoodsItems)
				{
                    //Click Add Button Goods Item
                    AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.AddGoodsItemBtn);

                    //Fill in Goods Item Name
                    AutomationHelpers.FillElement(driver, PurchaseRequestLocators.GoodsItemName, goodsItem.Name);

                    //Click Add Item
                    AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.AddGoodsBtn);

                    //Click Close Add Goods Item
                    AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.CloseGoodsBtn);

                    //Fill in the Quantity Goods Item
                    AutomationHelpers.FillElement(driver, PurchaseRequestLocators.QuantityGoods, goodsItem.Qty.ToString());

                    //kondisi jika pr type non catalog
                    if (param.PRType.Equals(PrTypeConstant.NonCatalog))
					{
                        //Edit Goods Items Price
                        AutomationHelpers.FillElement(driver, PurchaseRequestLocators.PriceGoodsItem, goodsItem.Price.ToString());
                    }
				}

                //Add Services Items
                foreach (PRItem servicesItem in param.ServicesItems)
                {
                    //Click Add Button Services Item
                    AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.AddServicesItemBtn);

                    //Fill in Services Item Name
                    AutomationHelpers.FillElement(driver, PurchaseRequestLocators.ServicesItemName, servicesItem.Name);

                    //Click Add Item
                    AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.AddServicesBtn);

                    //Click Close Add Services Item
                    AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.CloseServicesBtn);

                    //Fill in the Quantity Services Item
                    AutomationHelpers.FillElement(driver, PurchaseRequestLocators.QuantityServices, servicesItem.Qty.ToString());

                    //Condition if the PR Type (Non - Catalog)
                    if (param.PRType.Equals(PrTypeConstant.NonCatalog))
                    {
                        //Edit Goods Items Price
                        AutomationHelpers.FillElement(driver, PurchaseRequestLocators.PriceServicesItem, servicesItem.Price.ToString());
                    }
                }

                //Upload Attachment
                foreach (string path in param.AttachmentPaths)
                {
                    AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.AddAttachmentBtn);
                    AutomationHelpers.UploadFile(driver, PurchaseRequestLocators.AddFileBtnField, path);
                }

                if (!isResubmit)
				{
					if (isSubmit)
					{
                        //Click Submit Button
                        AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.SubmitBtn);

                        //Click Submit Confirmation Button
                        AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.YesBtn);
                    }
					else
					{
                        //Click Save as Draft Button
                        AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.SaveDraftBtn);

                        //Click Save as Draft Confirmation Button
                        AutomationHelpers.ClickElement(driver, PurchaseRequestLocators.YesBtn);
                    }

					string msg = AutomationHelpers.GetAlertMessage(driver, GlobalLocators.AlertSuccess);
					if (AutomationHelpers.ValidateAlert(msg, GlobalConfig.Config.SuccessMessages))
					{
						plan.Status = StatusConstant.Success;
						plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Successfully.";
						plan.RequestNumber = PageHelpers.SearchDocumentNumber(driver, ModuleNameConstant.PR, PurchaseRequestLocators.SearchList, PurchaseRequestLocators.RowListPrNumber, param.PRSubject);
					}
					else
					{
						throw new Exception(msg);
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
				if (isSubmit)
				{
					ExcelHelpers.WriteAutomationResult(cfg, plan);
					Thread.Sleep(GlobalConfig.Config.WaitWriteResultInSecond);
				}
			}

		}
	}
}
