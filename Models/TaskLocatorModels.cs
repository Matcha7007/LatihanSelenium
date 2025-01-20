using OpenQA.Selenium;

namespace LatihanSelenium.Models
{
	public class TaskLocatorModels
	{
		public By Search { get; set; } 
        public By RowTask { get; set; }
		public By ApprovalRemark { get; set; }
		public By BtnApprove { get; set; }
		public By BtnRevise { get; set; }
		public By BtnReject { get; set; }
		public By BtnResubmit { get; set; }
		public string Module { get; set; } = string.Empty;
		public string UrlPendingTask { get; set; } = string.Empty;
	}
}
