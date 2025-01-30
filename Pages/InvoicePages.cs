using LatihanSelenium.Constants;
using LatihanSelenium.Locators;
using LatihanSelenium.Models;
using LatihanSelenium.Utilities;
using OpenQA.Selenium;

namespace LatihanSelenium.Pages
{
    public static class InvoicePages
    {
        public static void HandleTestCase(IWebDriver driver, AppConfig cfg, TestplanModels plan, ResourceModels param)
        {
            try
            {
                switch (plan.TestCase)
                {
                    case TestCaseConstant.Submit:
                        InvoiceModels dataSubmit = param.Invoices.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataForConstant.Submit)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.Invoice, dataSubmit.Row);
                        HandleForm(driver, cfg, plan, dataSubmit, true);
                        break;
                    case TestCaseConstant.SaveAsDraft:
                        InvoiceModels dataDraft = param.Invoices.Where(x => x.TestCaseId.Equals(plan.TestCaseId) && x.DataFor.Equals(DataForConstant.SaveAsDraft)).FirstOrDefault()!;
                        plan.TestData = PlanHelper.CreateAddress(SheetConstant.Invoice, dataDraft.Row);
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

        public static void SearchListInvoice(IWebDriver driver)
        {
            try
            {
                InvoiceListModels param = new()
                {
                    SearchList = "IVCR/2024/09/0002"
                };

                AutomationHelpers.ElementExist(driver, DashboardLocators.Topbar);
                //Mengarahkan ke halaman PR
                driver.Navigate().GoToUrl(UrlConstant.CreateInvoice);

                //Input Search List
                AutomationHelpers.FillElement(driver, InvoiceLocators.SearchList, param.SearchList);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        public static void HandleForm(IWebDriver driver, AppConfig cfg, TestplanModels plan, InvoiceModels param, bool isSubmit = false, bool isResubmit = false)
        {
            try
            {
                if (!isResubmit)
                {
                    AutomationHelpers.NavigateTo(driver, UrlConstant.CreateInvoice);
                }

                //Click PO Number Lookup
                AutomationHelpers.ClickElement(driver, InvoiceLocators.LookupBtn);

                //Fill PO Number
                AutomationHelpers.FillElement(driver, InvoiceLocators.PONumber, param.PONumber);

                //Select PO
                AutomationHelpers.ClickElement(driver, InvoiceLocators.SelectPO);

                //Fill Vendor Invoice Number
                AutomationHelpers.FillElement(driver, InvoiceLocators.VendorInvoiceNumber, param.VendorInvoiceNumber);

                //Fill Invoice Received Date
                AutomationHelpers.FillDate(driver, InvoiceLocators.InvoiceReceivedDate, param.InvoiceReceivedDate);

                //Fill Invoice Original Date
                AutomationHelpers.FillDate(driver, InvoiceLocators.InvoiceOriginalDate, param.InvoiceOriginalDate);

                //Fill Invoice Due Date
                AutomationHelpers.FillDate(driver, InvoiceLocators.InvoiceDueDate, param.InvoiceDueDate);

                //Fill Payment Type
                AutomationHelpers.SelectElement(driver, InvoiceLocators.FieldPaymentType, InvoiceLocators.PaymentTypeList, param.PaymentType);

                if (param.PaymentType.Equals("Bank Transfer"))
                {
                    //Fill Bank Account
                    AutomationHelpers.SelectElement(driver, InvoiceLocators.FieldBankAccount, InvoiceLocators.BankAccountList, param.BankAccount);

                    //Fill Bank Vendor
                    AutomationHelpers.SelectElement(driver, InvoiceLocators.FieldBankVendor, InvoiceLocators.BankVendorList, param.BankVendor);
                }

                //Fill Remarks
                AutomationHelpers.FillElement(driver, InvoiceLocators.Remark, param.Remarks);

                //Click PO Drop Down
                AutomationHelpers.ClickElement(driver, InvoiceLocators.PODetailDropDown);

                //Add Goods VAT
                foreach (InvoiceGoodsItem goodsItem in param.GoodsItems)
                {

                    //Fill in Goods VAT Rate
                    AutomationHelpers.FillElement(driver, InvoiceLocators.GoodsVATRate, goodsItem.VATRate.ToString());
                }

                //Add Services Items
                foreach (InvoiceServicesItem servicesItem in param.ServicesItems)
                {

                    //Fill in Services VAT Rate
                    AutomationHelpers.FillElement(driver, InvoiceLocators.ServicesVATRate, servicesItem.VATRate.ToString());

                    //Fill in Services WHT Rate
                    AutomationHelpers.FillElement(driver, InvoiceLocators.ServicesWHTRate, servicesItem.WHTRate.ToString());

                    //Fill in Services WHT Amount
                    AutomationHelpers.FillElement(driver, InvoiceLocators.ServicesWHTAmount, servicesItem.WHTAmount.ToString());

                }

                //Fill Adjustment
                AutomationHelpers.FillElement(driver, InvoiceLocators.Adjustment, param.Adjustment);

                //Fill Adjustmant Value
                AutomationHelpers.FillElement(driver, InvoiceLocators.AdjustmentAmount, param.AdjustmentValue);

                //Upload Attachment
                AutomationHelpers.ClickElement(driver, InvoiceLocators.AttachmentSection);

                foreach (string path in param.AttachmentPaths)
                {
                    AutomationHelpers.UploadFile(driver, InvoiceLocators.AddAttachment, path);
                }

                if (!isResubmit)
                {
                    if (isSubmit)
                    {
                        //Click Submit Button
                        AutomationHelpers.ClickElement(driver, InvoiceLocators.submit);

                        //Click Submit Confirmation Button
                        AutomationHelpers.ClickElement(driver, InvoiceLocators.YesBtn);
                    }
                    else
                    {
                        //Click Save as Draft Button
                        AutomationHelpers.ClickElement(driver, InvoiceLocators.SaveAsDraft);

                        //Click Save as Draft Confirmation Button
                        AutomationHelpers.ClickElement(driver, InvoiceLocators.SaveAsDraftYesBtn);
                    }

                    string msg = AutomationHelpers.GetAlertMessage(driver, GlobalLocators.AlertSuccess);
                    if (AutomationHelpers.ValidateAlert(msg, GlobalConfig.Config.SuccessMessages))
                    {
                        plan.Status = StatusConstant.Success;
                        plan.Remarks = $"{plan.TestCase} {plan.ModuleName} is Successfully.";
                        plan.RequestNumber = PageHelpers.SearchDocumentNumber(driver, ModuleNameConstant.Invoice, InvoiceLocators.SearchList, InvoiceLocators.InvoiceRowList, param.PONumber);
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