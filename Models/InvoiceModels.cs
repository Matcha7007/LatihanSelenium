using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LatihanSelenium.Models
{
    public class InvoiceListModels
    {
        public string SearchList { get; set; } = string.Empty;
    }

    public class InvoiceModels : ModuleBase
    {
        public string PONumber { get; set; } = string.Empty;
        public string VendorInvoiceNumber { get; set; } = string.Empty;
        public string InvoiceReceivedDate { get; set; } = string.Empty;
        public string InvoiceOriginalDate { get; set; } = string.Empty;
        public string InvoiceDueDate { get; set; } = string.Empty;
        public string PaymentType { get; set; } = string.Empty;
        public string BankAccount { get; set; } = string.Empty;
        public string BankVendor { get; set; } = string.Empty;
        public string Remarks { get; set; } = string.Empty;
        public virtual List<InvoiceGoodsItem> GoodsItems { get; set; } = [];
        public virtual List<InvoiceServicesItem> ServicesItems { get; set; } = [];
        public string Adjustment { get; set; } = string.Empty;
        public string AdjustmentValue { get; set; } = string.Empty;
        public List<string> AttachmentPaths { get; set; } = [];
    }

    public class InvoiceGoodsItem
    {
        public int VATRate { get; set; }
    }

    public class InvoiceServicesItem
    {
        public int VATRate { get; set; }
        public int WHTRate { get; set; }
        public int WHTAmount { get; set; }
    }
}