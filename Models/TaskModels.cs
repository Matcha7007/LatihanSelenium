namespace LatihanSelenium.Models
{
	public class TaskModels : ModelBase
	{
		public string Action { get; set; } = string.Empty;
		public string Actor { get; set; } = string.Empty;
		public string DataNumber { get; set; } = string.Empty;
		public string ReturnTo { get; set; } = string.Empty;
		public string Notes { get; set; } = string.Empty;
		public string Result { get; set; } = string.Empty;
		public string Screenshot { get; set; } = string.Empty;
	}
}
