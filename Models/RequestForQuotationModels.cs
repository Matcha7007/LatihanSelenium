namespace LatihanSelenium.Models
{
    public class RequestForQuotationListModels
    {
        public string SearchList { get; set; } = string.Empty;
    }

    public class RequestForQuotationModels : ModuleBase
    {
        public string RFQSubject { get; set; } = string.Empty;
        public string RFQType { get; set; } = string.Empty;
        public string TenderBriefingDate { get; set; } = string.Empty;
        public string QuotationSubmissionDate { get; set; } = string.Empty;
        public string TargetAppointmentDate { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
        public virtual List<RFQItem> GoodsItems { get; set; } = [];
        public virtual List<RFQItem> ServicesItems { get; set; } = [];
        public string Vendor { get; set; } = string.Empty;
        public List<string> InternalAttachmentPaths { get; set; } = [];
        public List<string> VendorAttachmentPaths { get; set; } = [];
    }

    public class RFQItem
    {
        public string Name { get; set; } = string.Empty;
        public string Remark { get; set; } = string.Empty;
    }

    public class RFQVendor
    {
        public string Name { get; set; } = string.Empty;
    }
}