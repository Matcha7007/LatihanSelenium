using OpenQA.Selenium;

namespace LatihanSelenium.Locators
{
    public class VendorQuotationLocators
    {
        public static By SearchList = By.XPath("//div[@id='rc-tabs-2-panel-list']//input[@placeholder='Search']");
        public static By RowQuotationNumber = By.XPath("//tbody/tr[data-row-key]/td[1]/div[@class='ellipsis']");

        // Create VQ Header
        public static By UploadFile = By.XPath("//div[@class='wrap-input-upload-file']//label[@class='label']");
        public static By RFQNumberField = By.XPath("//div[@class='vendor-quotation-section-header']//button[@type='button']");
        public static By FillRFQNumber = By.XPath("//div[@class=' custom-form-item-filter']//input[@id='rfqNo']");
        public static By SelectRFQNumber = By.XPath("//td[@class='ant-table-cell']//button[@type='button']");
        public static By VendorNameField = By.XPath("//input[@id='form-vendor-quotation_vendorName']");
        public static By VendorNameList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By QuotationNumber = By.XPath("//input[@id='form-vendor-quotation_quotationNo']");
        public static By QuotationDate = By.XPath("//input[@id='form-vendor-quotation_quotationDate']");
        public static By ValidDate = By.XPath("//input[@id='form-vendor-quotation_validUntil']");
        public static By RemarksVQ = By.XPath("//textarea[@id='form-vendor-quotation_remarks']");

        //input[@id='form-vendor-quotation_goods_price_0
        // Items Section (Goods Item)
        public static By ItemBtn = By.XPath("//div[@class='ant-collapse-item vendor-quotation-form-collapse-panel-items']");
        public static By RemarksVendorGoodsItem = By.XPath("//textarea[@id='form-vendor-quotation_goods_remarks_from_vendor_0']");
        public static By PriceVendorGoodsItem = By.XPath("//div[@class='ant-input-number-input-wrap']//input']");
        public static By DiscountVendorGoodsItem = By.XPath("//input[@id='form-vendor-quotation_goods_discount_0']");
        public static By VendorGoodsItemVATRate = By.XPath("//input[@id='form-vendor-quotation_goodsTax']");

        // Items Section (Services Item)
        public static By RemarksVendorServicesItem = By.XPath("//textarea[@id='form-vendor-quotation_service_remarks_from_vendor_0']");
        public static By PriceVendorServicesItem = By.XPath("//input[@id='form-vendor-quotation_service_price_0']");
        public static By DiscountVendorServicesItem = By.XPath("//input[@id='form-vendor-quotation_service_discount_0']");
        public static By VendorServicesItemVATRate = By.XPath("//input[@id='form-vendor-quotation_servicesTax']");

        // Attachment
        public static By AttachmentBtn = By.XPath("(//div[@class='ant-collapse-header'])[3]");
        public static By AddAttachmentFileBtn = By.XPath("//div[@class='vendor-quotation-section-attachments']//div[@class='wrap-input-upload-file']");

        // Submit Button
        public static By SubmitBtn = By.XPath("//button[@type='submit']");
        public static By ConfirmSubmitBtn = By.XPath("(//div[@id='swal2-html-container']//button[@type='button'])[2]");
    }
}
