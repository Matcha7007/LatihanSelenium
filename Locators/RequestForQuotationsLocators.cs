using OpenQA.Selenium;
using static OpenQA.Selenium.BiDi.Modules.BrowsingContext.Locator;

namespace LatihanSelenium.Locators
{
    public class RequestForQuotationLocators
    {
        public static By SearchList = By.XPath("//div[@id='rc-tabs-4-panel-list']//input[@placeholder='Search']");
        public static By FieldRFQList = By.XPath("//input[@id='form-rfq_rfqTypeId']");
        public static By RFQList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By RFQSubject = By.XPath("//div[@class='ant-form-item-control-input-content']//input[@class='ant-input css-j3ybdg']");
        public static By TenderBriefingDate = By.XPath("//input[@id='form-rfq_tenderBriefingDate']");
        public static By QuotationDueDate = By.XPath("//input[@id='form-rfq_quotationSubmissionDueDate']");
        public static By TargetAppointmentDate = By.XPath("//input[@id='form-rfq_targetAppointmentDate']");
        public static By Remarks = By.XPath("//div[@class='ant-form-item-control-input-content']//textarea[@id='form-rfq_remarks']");
        public static By RowListRFQNumber = By.XPath("//tbody/tr[@data-row-key]/td[1]//div[@class='ellipsis']");

        //Items
        public static By ItemDropDown = By.XPath("//div[@class='ant-collapse-item rfq-form-collapse-panel-items']");

        //Goods Item
        public static By GoodsItemBtn = By.XPath("//div[@class='ant-space css-j3ybdg ant-space-horizontal ant-space-align-center mb-1']//button[@type='button']");
        public static By Goods_PRNumber = By.XPath("//input[@id='prNo']");
        public static By Goods_PRDateFrom = By.XPath("//input[@placeholder='PR Date From']");
        public static By Goods_PRDateTo = By.XPath("//input[@placeholder='PR Date To']");
        public static By Goods_PRSubject = By.XPath("//input[@id='prSubject']");
        public static By Goods_ItemName = By.XPath("//input[@id='itemName']");
        public static By FieldGoods_ItemCategory = By.XPath("//input[@id='itemCategoryId']");
        public static By Goods_ItemCategoryList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By AddGoodsItemBtn = By.XPath("//tbody[@class='ant-table-tbody']//button[@type='button']");
        public static By CloseGoodsItem = By.XPath("//div[@class='ant-modal-content']//button[@type='button']");
        public static By GoodsItemRemarks = By.XPath("//textarea[@id='form-rfq_goods_remarks_0']");

        //Services Item
        public static By ServicesItemBtn = By.XPath("//div[@class='ant-space css-j3ybdg ant-space-horizontal ant-space-align-center mb-1']//button[@type='button']");
        public static By Services_PRNumber = By.XPath("//input[@id='prNo']");
        public static By Services_PRDateFrom = By.XPath("//input[@placeholder='PR Date From']");
        public static By Services_PRDateTo = By.XPath("//input[@placeholder='PR Date To']");
        public static By Services_PRSubject = By.XPath("//input[@id='prSubject']");
        public static By Services_ItemName = By.XPath("//input[@id='itemName']");
        public static By FieldServices_ItemCategory = By.XPath("//input[@id='itemCategoryId']");
        public static By Services_ItemCategoryList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By AddServicesItemBtn = By.XPath("//tbody[@class='ant-table-tbody']//td[@class='ant-table-cell']//button[@type='button']");
        public static By CloseServicesItem = By.XPath("//div[@class='ant-modal-content']//button[@type='button']");
        public static By ServicesItemRemarks = By.XPath("//textarea[@id='form-rfq_service_remarks_0']");

        //Vendor
        public static By VendorDropDown = By.XPath("//div[@class='ant-collapse-item']");

        public static By VendorBtn = By.XPath("//div[@class='rfq-section-vendor']//button[@type='button']");
        public static By VendorName = By.XPath("//input[@id='vendorName']");
        public static By FieldIndustryType = By.XPath("//input[@id='vendorIndustryType']");
        public static By IndustryTypeList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By FieldPKPStatus = By.XPath("//input[@id='pkpStatus']");
        public static By PKPStatusList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By FieldVendorSize = By.XPath("//input[@id='vendorSize']");
        public static By VendorSizeList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By AddVendor = By.XPath("//tbody[@class='ant-table-tbody']//td[@class='ant-table-cell']//button[@type='button']");

        public static By AttachmentSection = By.XPath("//div[@class='ant-collapse-item rfq-form-collapse-panel-notes-and-attachments']");
        public static By BtnAddInternal = By.XPath("//div[@class='section-internal']//div[@class='wrap-input-upload-file']");
        public static By BtnAddVendor = By.XPath("//div[@class='section-vendor']//div[@class='wrap-input-upload-file']");

        public static By submit = By.XPath("//button[@type='submit']");
        public static By YesBtn = By.XPath("//div[@class='custom-alert-footer']//button[@class='ant-btn css-14wwjjs ant-btn-primary  button-custom']");
        public static By SaveAsDraft = By.XPath("//button[@class='ant-btn css-j3ybdg ant-btn-primary  button-custom']");
        public static By SaveAsDraftYesBtn = By.XPath("//div[@class='custom-alert-footer']//button[@class='ant-btn css-14wwjjs ant-btn-primary  button-custom']");

    }  

    public class RequestForQuotationTaskLocators
    {
        public static By Search = By.XPath("//div[@class='wrapper-rfq-pending-task-list']//input[@placeholder='Search']");
        public static By RowTask = By.XPath("//div[@class='wrapper-rfq-pending-task-list']//tbody/tr[@data-row-key]/td[1]//div[@class='ellipsis']");
        public static By ApprovalRemark = By.XPath("//textarea[@id='form-rfq_remarksApproval']");
        public static By BtnApprove = By.XPath("//button[@class='ant-btn css-j3ybdg ant-btn-primary  button-custom']");
        public static By BtnRevise = By.XPath("//button[@class='ant - btn css - j3ybdg ant - btn - dashed ant - btn - dangerous  button - custom']");
        public static By BtnReject = By.XPath("//button[@class='ant-btn css-j3ybdg ant-btn-primary ant-btn-dangerous  button-custom']");
        public static By BtnRsubmit = By.XPath("//form[@id='form-revise-rfq']//button[@type='submit']");
    }
}
