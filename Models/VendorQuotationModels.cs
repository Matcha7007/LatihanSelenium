using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LatihanSelenium.Models
{
    public class VendorQuotationListModels
    {
        public string SearchList { get; set; } = string.Empty;
    }

    public class VendorQuotationModels: ModuleBase
    {
        public string RFQNumber { get; set; } = string.Empty;

        public string VendorName { get; set; } = string.Empty;

        public string QuotationNumber { get; set; } = string.Empty;

        public string QuotationDate { get; set; } = string.Empty;

        public string ValidDate { get; set; } = string.Empty;

        public string RemarksVQ { get; set; } = string.Empty;

        public virtual List<VQItem> GoodsItem { get; set; } = [];

        public virtual List<VQItem> ServicesItem { get; set; } = [];

        public int GoodsVATRate { get; set; }

        public int ServicesVATRate { get; set; }

        public List<string> AttachmentPaths { get; set; } = [];
    }

    public class VQItem
    {
        public string Name { get; set; } = string.Empty;
        public string RemarksVendor { get; set; } = string.Empty;

        public int Price { get; set; }

        public int Discount { get; set; }
    }
}
