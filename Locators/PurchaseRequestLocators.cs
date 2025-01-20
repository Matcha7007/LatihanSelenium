using OpenQA.Selenium;

namespace LatihanSelenium.Locators
{
	public class PurchaseRequestLocators
	{
		public static By SearchList = By.XPath("//div[@id='rc-tabs-0-panel-list']//input[@placeholder='Search']");
		public static By RowListPrNumber = By.XPath("//tbody/tr[@data-row-key]/td[1]//div[@class='ellipsis']");

        // public static By FieldPRList = By.XPath("//input[@id='form-create-pr_prTypeId']");
        // public static By PRList = By.XPath("//div[@class='ant-select-item-option-content']");

        // public static By AttachmentSection = By.XPath("//form//div[@class='ant-collapse-item']");
        // public static By BtnAdd = By.XPath("//div[@class='purchase-request-section-attachments']//div[@class='wrap-input-upload-file']");

        // Create PR Header
        public static By FillSubjectPR = By.XPath("//form[@id='form-create-pr']//input[@id='form-create-pr_prSubject']");
        public static By CostCenterPRField = By.XPath("//input[@id='form-create-pr_costCenter']");
        public static By CostCenterPRList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By TypePRField = By.XPath("//input[@id='form-create-pr_prTypeId']");
        public static By TypePRList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By RequiredDatePR = By.XPath("//input[@id='form-create-pr_requiredDate']");
        public static By ProjectCodePR = By.XPath("//input[@id='form-create-pr_projectCode']");
        public static By RemarksPR = By.XPath("//textarea[@id='form-create-pr_remarks']");

        // Create PR Items (Goods Item)
        public static By ItemBtn = By.XPath("//div[@class='ant-collapse-item pr-form-collapse-panel-items']");
        public static By AddGoodsItemBtn = By.XPath("//button[@type='button']");
        public static By GoodsItemCode = By.XPath("//input[@id='itemCode']");
        public static By GoodsItemName = By.XPath("//input[@id='itemName']");
        public static By GoodsItemCategoryField = By.XPath("//input[@id='itemCategoryId']");
        public static By GoodsItemCategoryList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By AddGoodsBtn = By.XPath("//tbody[@class='ant-table-tbody']//button[@type='button']");
        public static By CloseGoodsBtn = By.XPath("//button[@aria-label='Close']");
        public static By QuantityGoods = By.XPath("//input[@id='form-create-pr_goods_quantity_0']");
        public static By PriceGoodsItem = By.XPath("//input[@id='form-create-pr_goods_price_0']");

        // Create PR Items (Services Item)
        public static By AddServicesItemBtn = By.XPath("(//button[@type='button'])[4]");
        public static By ServicesItemCode = By.XPath("//input[@id='itemCode']");
        public static By ServicesItemName = By.XPath("//input[@id='itemName']");
        public static By ServicesItemCategoryField = By.XPath("//input[@id='itemCategoryId']");
        public static By ServicesItemCategoryList = By.XPath("//div[@class='ant-select-item-option-content']");
        public static By AddServicesBtn = By.XPath("//td[@class='ant-table-cell']//button[@type='button']");
        public static By CloseServicesBtn = By.XPath("//button[@aria-label='Close']");
        public static By QuantityServices = By.XPath("//input[@id='form-create-pr_services_quantity_0']");
        public static By PriceServicesItem = By.XPath("//input[@id='form-create-pr_services_price_0']");

        // Add Attachments
        public static By AddAttachmentBtn = By.XPath("(//div[@class='ant-collapse-header'])[3]");
        public static By AddFileBtnField = By.XPath("//div[@class='purchase-request-section-attachments']//div[@class='wrap-input-upload-file']");
        public static By FileAttachment = By.XPath("//input[@type='file']");

        // Submit & Save as Draft Button
        public static By SubmitBtn = By.XPath("//button[@type='submit']");
        public static By YesBtn = By.XPath("(//div[@id='swal2-html-container']//button[@type='button'])[2]");
        public static By SaveDraftBtn = By.XPath("");

        //public static By AttachmentSection = By.XPath("//form//div[@class='ant-collapse-item']");
        //public static By BtnAdd = By.XPath("//div[@class='purchase-request-section-attachments']//div[@class='wrap-input-upload-file']");
    }

    public class PurchaseRequestTaskLocators
	{
		public static By Search = By.XPath("//div[@class='wrapper-pr-pending-task-list']//input[@placeholder='Search']");
		public static By RowTask = By.XPath("//div[@class='wrapper-pr-pending-task-list']//tbody/tr[@data-row-key]/td[1]//div[@class='ellipsis']");
		public static By ApprovalRemark = By.XPath("//textarea[@id='form-approval-pr_remarksApproval']");
		public static By BtnApprove = By.XPath("(//form[@id='form-approval-pr']//button[@type='button'])[2]");
		public static By BtnRevise = By.XPath("(//form[@id='form-approval-pr']//button[@type='button'])[3]");
		public static By BtnReject = By.XPath("(//form[@id='form-approval-pr']//button[@type='button'])[4]");
		public static By BtnRsubmit = By.XPath("//form[@id='form-revise-pr']//button[@type='submit']");
	}
}
