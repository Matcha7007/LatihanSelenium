using LatihanSelenium.Constants;
using LatihanSelenium.Locators;
using LatihanSelenium.Models;
using LatihanSelenium.Utilities;
using OpenQA.Selenium;

namespace LatihanSelenium.Pages
{
	public static class RequestForQuotationPages
	{
        public static void HandleTestCase(IWebDriver driver, AppConfig cfg, TestplanModels plan, ResourceModels param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Submit:
                        RequestForQuotationModels dataSubmit = param.RequestForQuotations.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataForConstant.Submit)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.RFQ, dataSubmit.Row);
                        HandleForm(driver, cfg, plan, dataSubmit, true);
                        break;
                    case TestCaseConstant.SaveAsDraft:
                        RequestForQuotationModels dataDraft = param.RequestForQuotations.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataForConstant.SaveAsDraft)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.RFQ, dataDraft.Row);
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

        public static void SearchListRFQ(IWebDriver driver)
		{
			try
			{
				PurchaseRequestListModels param = new()
				{
					SearchList = "RFQ/2024/09/0001"
                };

				AutomationHelpers.ElementExist(driver, DashboardLocators.Topbar);
				//Mengarahkan ke halaman RFQ
				driver.Navigate().GoToUrl(UrlConstant.CreateRFQ);

				//Input Search List
				AutomationHelpers.FillElement(driver, RequestForQuotationLocators.SearchList, param.SearchList);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
			}
		}

		public static void HandleForm(IWebDriver driver, AppConfig cfg, TestplanModels plan, RequestForQuotationModels param, bool isSubmit = false, bool isResubmit = false)
		{
			try
			{
				if (!isResubmit)
				{
					AutomationHelpers.NavigateTo(driver, UrlConstant.CreateRFQ);
				}

                //Create RFQ Header Section
                //Fill in the RFQ Subject
                AutomationHelpers.FillElement(driver, RequestForQuotationLocators.RFQSubject, param.RFQSubject);

                //Select RQF Type
                AutomationHelpers.SelectElement(driver, RequestForQuotationLocators.FieldRFQList, RequestForQuotationLocators.RFQList, param.RFQType);

                //Fill in Tender Briefing Date
                AutomationHelpers.FillDate(driver, RequestForQuotationLocators.TenderBriefingDate, param.TenderBriefingDate);

                //Fill in Quotation Submission Date
                AutomationHelpers.FillDate(driver, RequestForQuotationLocators.QuotationDueDate, param.QuotationSubmissionDate);

                //Fill in Target Appointment Date
                AutomationHelpers.FillDate(driver, RequestForQuotationLocators.TargetAppointmentDate, param.TargetAppointmentDate);

                //Fill in Remarks
                AutomationHelpers.FillElementNonMandatory(driver, RequestForQuotationLocators.Remarks, param.Remarks);

                //Create RFQ Items Section
                //buka item section
                AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.ItemDropDown);

                //Add Goods Items
                foreach (RFQItem goodsItem in param.GoodsItems)
				{
                    //Click Add Button Goods Item
                    AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.GoodsItemBtn);

                    //Fill in Goods Item Name
                    AutomationHelpers.FillElement(driver, RequestForQuotationLocators.Goods_ItemName, goodsItem.Name);

                    //Click Add Item
                    AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.AddGoodsItemBtn);

                    //Click Close Add Goods Item
                    AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.CloseGoodsItem);

                    //Fill in the Remarks Goods Item
                    AutomationHelpers.FillElement(driver, RequestForQuotationLocators.GoodsItemRemarks, goodsItem.Remark);
				}

                //Add Services Items
                foreach (RFQItem servicesItem in param.ServicesItems)
                {
                    //Click Add Button Services Item
                    AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.ServicesItemBtn);

                    //Fill in Services Item Name
                    AutomationHelpers.FillElement(driver, RequestForQuotationLocators.Services_ItemName, servicesItem.Name);

                    //Click Add Item
                    AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.AddServicesItemBtn);

                    //Click Close Add Services Item
                    AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.CloseServicesItem);

                    //Fill in the Quantity Services Item
                    AutomationHelpers.FillElement(driver, RequestForQuotationLocators.ServicesItemRemarks, servicesItem.Remark);

                }

                //Create RFQ Vendor Section
                //buka Vendor section
                AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.VendorDropDown);

                //Click Add Button Vendor
                AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.VendorBtn);

                //Fill in Vendor Name
                AutomationHelpers.FillElement(driver, RequestForQuotationLocators.VendorName, param.Vendor);

                //Click Add Item
                AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.AddVendor);

                //Click Close 
                AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.CloseVendor);


                //Upload Internal Attachment
                foreach (string path in param.InternalAttachmentPaths)
                {
                    AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.AttachmentSection);
                    AutomationHelpers.UploadFile(driver, RequestForQuotationLocators.BtnAddInternal, path);
                }

                //Upload Vendor Attachment
                foreach (string path in param.InternalAttachmentPaths)
                {
                    AutomationHelpers.UploadFile(driver, RequestForQuotationLocators.BtnAddVendor, path);
                }

                if (!isResubmit)
				{
					if (isSubmit)
					{
                        //Click Submit Button
                        AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.submit);

                        //Click Submit Confirmation Button
                        AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.YesBtn);
                    }
					else
					{
                        //Click Save as Draft Button
                        AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.SaveAsDraft);

                        //Click Save as Draft Confirmation Button
                        AutomationHelpers.ClickElement(driver, RequestForQuotationLocators.SaveAsDraftYesBtn);
                    }

					string msg = AutomationHelpers.GetAlertMessage(driver, GlobalLocators.AlertSuccess);
					if (AutomationHelpers.ValidateAlert(msg, GlobalConfig.Config.SuccessMessages))
					{
						plan.Status = StatusConstant.Success;
						plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Successfully.";
						plan.RequestNumber = PageHelpers.SearchDocumentNumber(driver, ModuleNameConstant.RFQ, RequestForQuotationLocators.SearchList, RequestForQuotationLocators.RowListRFQNumber, param.RFQSubject);
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
