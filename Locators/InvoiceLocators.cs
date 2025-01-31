using OpenQA.Selenium;

namespace LatihanSelenium.Locators
{
    public class InvoiceLocators
    {
        public static By SearchList = By.XPath("//div[@id='rc-tabs-3-panel-list']//input[@placeholder='Search']");
        public static By InvoiceRowList = By.XPath("//tbody/tr[@data-row-key]/td[1]//div[@class='ellipsis']");
        public static By LookupBtn = By.XPath("//span[@class='ant-input-wrapper ant-input-group css-j3ybdg']//button[@type='button']");
        public static By PONumber = By.XPath("//div[@class='ant-modal-content']//input[@id='poNo']");
        public static By SelectPO = By.XPath("//div[@class='ant-modal-content']//td[@class='ant-table-cell']//button[@type='button']");
        public static By ClosePO = By.XPath("//div[@class='ant-modal-content']//button[@aria-label='Close']");
        public static By VendorInvoiceNumber = By.XPath("//form//input[@id='form-create-invoice-receive_vendorInvoiceNo']");
        public static By InvoiceReceivedDate = By.XPath("//form//input[@id='form-create-invoice-receive_invoiceReceivedDate']");
        public static By InvoiceOriginalDate = By.XPath("//form//input[@id='form-create-invoice-receive_originalInvoiceDate']");
        public static By InvoiceDueDate = By.XPath("//form//input[@id='form-create-invoice-receive_invoiceDueDate']");
        public static By FieldPaymentType = By.XPath("//input[@id='form-create-invoice-receive_paymentTypeId']");
        public static By PaymentTypeList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By FieldBankAccount = By.XPath("//input[@id='form-create-invoice-receive_bankAccountId']");
        public static By BankAccountList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By FieldBankVendor = By.XPath("//input[@id='form-create-invoice-receive_vendorBankAccountId']");
        public static By BankVendorList = By.XPath("//div[@class='ant-select-item-option-content']");

        public static By Remark = By.XPath("//form//textarea[@id='form-create-invoice-receive_remarks']");

        public static By PODetailDropDown = By.XPath("//div[@class='ant-collapse-item invoice-receive-form-collapse-panel-items']");
        public static By GoodsVATRate = By.XPath("//input[@class='ant-input-number-input']");

        public static By ServicesVATRate = By.XPath("//input[@class='ant-input-number-input']");
        public static By ServicesWHTRate = By.XPath("//input[@class='ant-input-number-input']");
        public static By ServicesWHTAmount = By.XPath("//input[@class='ant-input-number-input']");

        public static By Adjustment = By.XPath("//textarea[@id='form-create-invoice-receive_adjustmentReason']");
        public static By AdjustmentAmount = By.XPath("//input[@id='form-create-invoice-receive_adjustmentAmount']");

        public static By AttachmentSection = By.XPath("//div[@class='ant-collapse-item']");
        public static By AddAttachment = By.XPath("//div[@class='invoice-receive-section-attachments']//div[@class='wrap-input-upload-file']");

        public static By submit = By.XPath("//button[@type='submit']");
        public static By YesBtn = By.XPath("//div[@class='custom-alert-footer']//button[@class='ant-btn css-14wwjjs ant-btn-primary  button-custom']");
        public static By SaveAsDraft = By.XPath("//button[@class='ant-btn css-j3ybdg ant-btn-primary  button-custom']");
        public static By SaveAsDraftYesBtn = By.XPath("//div[@class='custom-alert-footer']//button[@class='ant-btn css-14wwjjs ant-btn-primary  button-custom']");
    }

    public class InvoiceTaskLocators
    {
        public static By Search = By.XPath("//input[@placeholder='Search']");
        public static By RowTask = By.XPath("//div[@class='wrapper-invoice-pending-task-list']//tbody/tr[@data-row-key]/td[1]//div[@class='ellipsis']");
        public static By ApprovalRemark = By.XPath("//textarea[@id='form-create-invoice-receive_remarksApproval']");
        public static By BtnApprove = By.XPath("//button[@class='ant-btn css-j3ybdg ant-btn-primary  button-custom']");
        public static By BtnRevise = By.XPath("//button[@class='ant - btn css - j3ybdg ant - btn - dashed ant - btn - dangerous  button - custom']");
        public static By BtnReject = By.XPath("//button[@class='ant-btn css-j3ybdg ant-btn-primary ant-btn-dangerous  button-custom']");
        public static By BtnRsubmit = By.XPath("//form[@id='form-revise-invoice']//button[@type='submit']");
    }
}