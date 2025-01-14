namespace LatihanSelenium.Models
{
	public class AppConfig
	{
		public string Url { get; set; } = string.Empty;
		public string LogPath { get; set; } = string.Empty;
		public string ExcelPath { get; set; } = string.Empty;
		public string FileTestPath { get; set; } = string.Empty;
		public string ScreenCapturePath { get; set; } = string.Empty;
		public string TestData { get; set; } = string.Empty;
		public int WaitElementInSecond { get; set; }
		public int WaitWriteResultInSecond { get; set; }
		public int MaxRetrySearchTask { get; set; }
		public int WaitRetrySearchTaskInMiliSecond { get; set; }
		public string[] ErrorMessages { get; set; } = [];
		public string[] SuccessMessages { get; set; } = [];
	}

	public static class GlobalConfig
	{
		public static AppConfig Config { get; set; } = new();
	}
}
