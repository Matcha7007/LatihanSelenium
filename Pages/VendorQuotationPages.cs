using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LatihanSelenium.Constants;
using LatihanSelenium.Locators;
using LatihanSelenium.Models;
using LatihanSelenium.Utilities;
using OpenQA.Selenium;

namespace LatihanSelenium.Pages
{
    public class VendorQuotationPages
    {
        public static void HandleTestCase(IWebDriver driver, AppConfig cfg, TestplanModels plan, ResourceModels param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Submit:
                        VendorQuotationModels dataSubmit = param.VendorQuotations.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataForConstant.Submit)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.VQ, dataSubmit.Row);
                        HandleForm(driver, cfg, plan, dataSubmit, true);
                        break;
                    default: break;
                }
            }
            finally
            {
                driver.Dispose();
            }
        }

        public static void SearchListVQ(IWebDriver driver)
        {
            try
            {
                VendorQuotationListModels param = new()
                {
                    SearchList = "QN/2024-03-27/003"
                };

                AutomationHelpers.ElementExist(driver, DashboardLocators.Topbar);

                //Mengarahkan ke halaman VQ
                driver.Navigate().GoToUrl(UrlConstant.CreateVQ);

                //Input Search List
                AutomationHelpers.FillElement(driver, RequestForQuotationLocators.SearchList, param.SearchList);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        public static void HandleForm(IWebDriver driver, AppConfig cfg, TestplanModels plan, VendorQuotationModels param, bool isSubmit = false, bool isResubmit = false)
        {
            try
            {
                if (!isResubmit)
                {
                    AutomationHelpers.NavigateTo(driver, UrlConstant.CreateVQ);
                }

                // File Upload
                // AutomationHelpers.ClickElement(driver, VendorQuotationLocators.UploadFile);

                // Select RFQ Number
                AutomationHelpers.ClickElement(driver, VendorQuotationLocators.RFQNumberField);
                AutomationHelpers.FillElement(driver, VendorQuotationLocators.FillRFQNumber, param.RFQNumber);
                AutomationHelpers.ClickElement(driver, VendorQuotationLocators.SelectRFQNumber);

                // Select Vendor Name
                AutomationHelpers.SelectElement(driver, VendorQuotationLocators.VendorNameField, VendorQuotationLocators.VendorNameList, param.VendorName);

                // Fill in Quotation Number
                AutomationHelpers.FillElement(driver, VendorQuotationLocators.QuotationNumber, param.QuotationNumber);

                // Fill in Quotation Date
                AutomationHelpers.FillDate(driver, VendorQuotationLocators.QuotationDate, param.QuotationDate);

                // Fill in Valid Date
                AutomationHelpers.FillDate(driver, VendorQuotationLocators.ValidDate, param.ValidDate);

                // Fill in Remarks
                AutomationHelpers.FillElementNonMandatory(driver, VendorQuotationLocators.RemarksVQ, param.RemarksVQ);

                // Open Items Section
                AutomationHelpers.ClickElement(driver, VendorQuotationLocators.ItemBtn);

                foreach (VQItem goodsItem in param.GoodsItem)
                {
                    // Fill in Remarks Vendor
                    AutomationHelpers.FillElementNonMandatory(driver, VendorQuotationLocators.RemarksVendorGoodsItem, goodsItem.RemarksVendor);

                    // Fill in Price
                    AutomationHelpers.FillElement(driver, VendorQuotationLocators.PriceVendorGoodsItem, goodsItem.Price.ToString());

                    // Fill in Discount
                    AutomationHelpers.FillElement(driver, VendorQuotationLocators.DiscountVendorGoodsItem, goodsItem.Discount.ToString());
                }

                foreach (VQItem servicesItem in param.ServicesItem)
                {
                    // Fill in Remarks Vendor
                    AutomationHelpers.FillElementNonMandatory(driver, VendorQuotationLocators.RemarksVendorServicesItem, servicesItem.RemarksVendor);

                    // Fill in Price
                    AutomationHelpers.FillElement(driver, VendorQuotationLocators.PriceVendorServicesItem, servicesItem.Price.ToString());

                    // Fill in Discount
                    AutomationHelpers.FillElement(driver, VendorQuotationLocators.DiscountVendorServicesItem, servicesItem.Discount.ToString());
                }

                // Fill in Goods VAT Rate
                AutomationHelpers.FillElement(driver, VendorQuotationLocators.VendorGoodsItemVATRate, param.GoodsVATRate.ToString());

                // FIll in Services VAT Rate
                AutomationHelpers.FillElement(driver, VendorQuotationLocators.VendorServicesItemVATRate, param.ServicesVATRate.ToString());

                // Upload Attachment
                foreach (string path in param.AttachmentPaths)
                {
                    AutomationHelpers.ClickElement(driver, VendorQuotationLocators.AttachmentBtn);
                    AutomationHelpers.UploadFileNonMandatory(driver, VendorQuotationLocators.AddAttachmentFileBtn, path);
                }

                // Submit or Save as Draft
                if (!isResubmit)
                {
                    if (isSubmit)
                    {
                        // Click Submit Button
                        AutomationHelpers.ClickElement(driver, VendorQuotationLocators.SubmitBtn);

                        // Click Confirm Submit Button
                        AutomationHelpers.ClickElement(driver, VendorQuotationLocators.ConfirmSubmitBtn);
                    }

                    string msg = AutomationHelpers.GetAlertMessage(driver, GlobalLocators.AlertSuccess);
                    if (AutomationHelpers.ValidateAlert(msg, GlobalConfig.Config.SuccessMessages))
                    {
                        plan.Status = StatusConstant.Success;
                        plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Successfully.";
                        plan.RequestNumber = PageHelpers.SearchDocumentNumber(driver, ModuleNameConstant.VQ, VendorQuotationLocators.SearchList, VendorQuotationLocators.RowQuotationNumber, param.QuotationNumber);
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
