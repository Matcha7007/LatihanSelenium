using LatihanSelenium.Constants;
using LatihanSelenium.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;
using System.Runtime.InteropServices;

namespace LatihanSelenium.Utilities
{
	public static class PlanHelper
	{
		public static string CreateAddress(string address, int row)
		{
			return $"{address}!A{row + 1}";
		}
	}

	public static class ExcelHelpers
	{
		private static Dictionary<string, PurchaseRequestModels> _PR = [];

		public static ResourceModels ReadExcel(AppConfig cfg)
		{
			ResourceModels result = new();
			try
			{
				Console.WriteLine($"Starting Read Automation Config");

				using (FileStream fs = new(cfg.ExcelPath, FileMode.Open, FileAccess.Read))
				{
					using (IWorkbook wb = new XSSFWorkbook(fs))
					{
						result.Users = ReadUsers(wb);
						result.Testplans = ReadTestPlans(wb);
						var newTestPlans = result.Testplans.Where(x => string.IsNullOrEmpty(x.Status)).ToList();

						result.Tasks = ReadTask(wb, newTestPlans);
						result.PurchaseRequests = ReadPR(wb, newTestPlans);

						wb.Close();
						fs.Close();
					}
				}
				Console.WriteLine($"End ReadAutomationConfig");
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Fail ReadAutomationConfig : \n - {ex.Message}");
			}
		}

		#region Read Sheet
		private static List<PurchaseRequestModels> ReadPR(IWorkbook wb, List<TestplanModels> plans)
		{
			List<PurchaseRequestModels> result = [];
			plans = plans.Where(x => x.ModuleName == ModuleNameConstant.PR).ToList();
			try
			{
				string workSheetName = SheetConstant.PR;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexDataFor = colIndexId + 1;
				int colIndexSeq = colIndexDataFor + 1;
				int colIndexPRSubject = colIndexSeq + 1;
				int colIndexCostCenter = colIndexPRSubject + 1;
				int colIndexPRType = colIndexCostCenter + 1;
				int colIndexRequiredDate = colIndexPRType + 1;
				int colIndexProjectCode = colIndexRequiredDate + 1;
				int colIndexRemarks = colIndexProjectCode + 1;
				int colIndexGoodsItemName = colIndexRemarks + 1;
				int colIndexGoodsItemQty = colIndexGoodsItemName + 1;
				int colIndexGoodsItemPrice = colIndexGoodsItemQty + 1;
				int colIndexServicesItemName = colIndexGoodsItemPrice + 1;
				int colIndexServicesItemQty = colIndexServicesItemName + 1;
				int colIndexServicesItemPrice = colIndexServicesItemQty + 1;
				int colIndexGoodsVATRate = colIndexServicesItemPrice + 1;
				int colIndexServicesVATRate = colIndexGoodsVATRate + 1;
				int colIndexAttachmentPath = colIndexServicesVATRate + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderDataFor = firstRow.GetCell(colIndexDataFor);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderPRSubject = firstRow.GetCell(colIndexPRSubject);
				ICell cellHeaderCostCenter = firstRow.GetCell(colIndexCostCenter);
				ICell cellHeaderPRType = firstRow.GetCell(colIndexPRType);
				ICell cellHeaderRequiredDate = firstRow.GetCell(colIndexRequiredDate);
				ICell cellHeaderProjectCode = firstRow.GetCell(colIndexProjectCode);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);
				ICell cellHeaderGoodsItemName = firstRow.GetCell(colIndexGoodsItemName);
				ICell cellHeaderGoodsItemQty = firstRow.GetCell(colIndexGoodsItemQty);
				ICell cellHeaderGoodsItemPrice = firstRow.GetCell(colIndexGoodsItemPrice);
				ICell cellHeaderServicesItemName = firstRow.GetCell(colIndexServicesItemName);
				ICell cellHeaderServicesItemQty = firstRow.GetCell(colIndexServicesItemQty);
				ICell cellHeaderServicesItemPrice = firstRow.GetCell(colIndexServicesItemPrice);
				ICell cellHeaderGoodsVATRate = firstRow.GetCell(colIndexGoodsVATRate);
				ICell cellHeaderServicesVATRate = firstRow.GetCell(colIndexServicesVATRate);
				ICell cellHeaderAttachmentPath = firstRow.GetCell(colIndexAttachmentPath);

				ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				ValidateCellHeaderName(cellHeaderDataFor, workSheetName, colIndexDataFor + 1, "Data For");
				ValidateCellHeaderName(cellHeaderPRSubject, workSheetName, colIndexPRSubject + 1, "PR Subject");
				ValidateCellHeaderName(cellHeaderCostCenter, workSheetName, colIndexCostCenter + 1, "Cost Center");
				ValidateCellHeaderName(cellHeaderPRType, workSheetName, colIndexPRType + 1, "PR Type");
				ValidateCellHeaderName(cellHeaderRequiredDate, workSheetName, colIndexRequiredDate + 1, "Required Date");
				ValidateCellHeaderName(cellHeaderProjectCode, workSheetName, colIndexProjectCode + 1, "Project Code");
				ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				ValidateCellHeaderName(cellHeaderGoodsItemName, workSheetName, colIndexGoodsItemName + 1, "Goods Item Name");
				ValidateCellHeaderName(cellHeaderGoodsItemQty, workSheetName, colIndexGoodsItemQty + 1, "Goods Item Qty");
				ValidateCellHeaderName(cellHeaderGoodsItemPrice, workSheetName, colIndexGoodsItemPrice + 1, "Goods Item Price");
				ValidateCellHeaderName(cellHeaderServicesItemName, workSheetName, colIndexServicesItemName + 1, "Services Item Name");
				ValidateCellHeaderName(cellHeaderServicesItemQty, workSheetName, colIndexServicesItemQty + 1, "Services Item Qty");
				ValidateCellHeaderName(cellHeaderServicesItemPrice, workSheetName, colIndexServicesItemPrice + 1, "Services Item Price");
				ValidateCellHeaderName(cellHeaderGoodsVATRate, workSheetName, colIndexGoodsVATRate + 1, "Goods VAT Rate");
				ValidateCellHeaderName(cellHeaderServicesVATRate, workSheetName, colIndexServicesVATRate + 1, "Services VAT Rate");
				ValidateCellHeaderName(cellHeaderAttachmentPath, workSheetName, colIndexAttachmentPath + 1, "Attachment Path");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (IsRowEmpty(currentRow, colIndexAttachmentPath))
					{
						break;
					}

					#region Assign Param
					string idStr = TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					string seqStr = TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum);
					string dataFor = TryGetString(currentRow.GetCell(colIndexDataFor), workSheetName, cellHeaderDataFor.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(idStr))
					{
						PurchaseRequestModels param = new();
						int id = int.Parse(idStr);
						if (plans.Any(x => x.TestCaseId.Equals(id)))
						{
							string goodsItemName = TryGetString(currentRow.GetCell(colIndexGoodsItemName), workSheetName, cellHeaderGoodsItemName.StringCellValue, rowNum);
							PRItem goodsItem = new();
							if (!string.IsNullOrEmpty(goodsItemName))
							{
								goodsItem = new()
								{
									Name = goodsItemName,									
									Qty = IntegerParse(TryGetString(currentRow.GetCell(colIndexGoodsItemQty), workSheetName, cellHeaderGoodsItemQty.StringCellValue, rowNum)),
									Price = IntegerParse(TryGetString(currentRow.GetCell(colIndexGoodsItemPrice), workSheetName, cellHeaderGoodsItemPrice.StringCellValue, rowNum))
								};
							}

							string servicesItemName = TryGetString(currentRow.GetCell(colIndexServicesItemName), workSheetName, cellHeaderServicesItemName.StringCellValue, rowNum);
							PRItem servicesItem = new();
							if (!string.IsNullOrEmpty(servicesItemName))
							{
								servicesItem = new()
								{
									Name = servicesItemName,
									Qty = IntegerParse(TryGetString(currentRow.GetCell(colIndexServicesItemQty), workSheetName, cellHeaderServicesItemQty.StringCellValue, rowNum)),
									Price = IntegerParse(TryGetString(currentRow.GetCell(colIndexServicesItemPrice), workSheetName, cellHeaderServicesItemPrice.StringCellValue, rowNum))
								};
							}

							string attachmentPath = TryGetString(currentRow.GetCell(colIndexAttachmentPath), workSheetName, cellHeaderAttachmentPath.StringCellValue, rowNum);							

							if (_PR.TryGetValue($"{idStr}{dataFor}{seqStr}", out PurchaseRequestModels? value))
							{
								if (!string.IsNullOrEmpty(goodsItem.Name)) value.GoodsItems.Add(goodsItem);
								if (!string.IsNullOrEmpty(servicesItem.Name)) value.ServicesItems.Add(servicesItem);
								if (!string.IsNullOrEmpty(attachmentPath)) value.AttachmentPaths.Add(attachmentPath);
							}
							else
							{
								param.Row = rowNum;
								param.TestCaseId = id;
								param.DataFor = dataFor;
								param.Sequence = !string.IsNullOrEmpty(seqStr) ? int.Parse(seqStr) : 1;
								param.PRSubject = TryGetString(currentRow.GetCell(colIndexPRSubject), workSheetName, cellHeaderPRSubject.StringCellValue, rowNum);
								param.CostCenter = TryGetString(currentRow.GetCell(colIndexCostCenter), workSheetName, cellHeaderCostCenter.StringCellValue, rowNum);
								param.PRType = TryGetString(currentRow.GetCell(colIndexPRType), workSheetName, cellHeaderPRType.StringCellValue, rowNum);
								param.RequiredDate = TryGetString(currentRow.GetCell(colIndexRequiredDate), workSheetName, cellHeaderRequiredDate.StringCellValue, rowNum);
								param.ProjectCode = TryGetString(currentRow.GetCell(colIndexProjectCode), workSheetName, cellHeaderProjectCode.StringCellValue, rowNum);
								param.Remarks = TryGetString(currentRow.GetCell(colIndexRemarks), workSheetName, cellHeaderRemarks.StringCellValue, rowNum);
								param.GoodsVATRate = IntegerParse(TryGetString(currentRow.GetCell(colIndexGoodsVATRate), workSheetName, cellHeaderGoodsVATRate.StringCellValue, rowNum));
								param.ServicesVATRate = IntegerParse(TryGetString(currentRow.GetCell(colIndexServicesVATRate), workSheetName, cellHeaderServicesVATRate.StringCellValue, rowNum));
								if (!string.IsNullOrEmpty(goodsItem.Name)) param.GoodsItems.Add(goodsItem);
								if (!string.IsNullOrEmpty(servicesItem.Name)) param.ServicesItems.Add(servicesItem);
								if (!string.IsNullOrEmpty(attachmentPath)) param.AttachmentPaths.Add(attachmentPath);
								_PR.Add($"{idStr}{dataFor}{seqStr}", param);
							}
						}
					}
					#endregion
				}
				result.AddRange(_PR.Values);
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read Sheet {SheetConstant.PR} is Fail : {CheckErrorReadExcel(ex.Message)}");
			}
		}

		private static List<TaskModels> ReadTask(IWorkbook wb, List<TestplanModels> plans)
		{
			List<TaskModels> result = [];
			plans = plans.Where(x => x.TestCase == TestCaseConstant.Approval).ToList();
			try
			{
				string workSheetName = SheetConstant.Task;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexSeq = colIndexId + 1;
				int colIndexActor = colIndexSeq + 1;
				int colIndexAction = colIndexActor + 1;
				int colIndexNotes = colIndexAction + 1;
				int colIndexResult = colIndexNotes + 1;
				int colIndexScreenshot = colIndexResult + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderSeq = firstRow.GetCell(colIndexSeq);
				ICell cellHeaderActor = firstRow.GetCell(colIndexActor);
				ICell cellHeaderAction = firstRow.GetCell(colIndexAction);
				ICell cellHeaderNotes = firstRow.GetCell(colIndexNotes);
				ICell cellHeaderResult = firstRow.GetCell(colIndexResult);

				ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Test Case Id");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (IsRowEmpty(currentRow, colIndexScreenshot))
					{
						break;
					}

					#region Assign Param
					string id = TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum);
					if (!string.IsNullOrEmpty(id))
					{
						string colResult = TryGetString(currentRow.GetCell(colIndexResult), workSheetName, cellHeaderResult.StringCellValue, rowNum);
						if (string.IsNullOrEmpty(colResult))
						{
							TaskModels param = new()
							{
								Row = rowNum,
								TestCaseId = int.Parse(id),
								Sequence = int.Parse(TryGetString(currentRow.GetCell(colIndexSeq), workSheetName, cellHeaderSeq.StringCellValue, rowNum)),
								Actor = TryGetString(currentRow.GetCell(colIndexActor), workSheetName, cellHeaderActor.StringCellValue, rowNum),
								Action = TryGetString(currentRow.GetCell(colIndexAction), workSheetName, cellHeaderAction.StringCellValue, rowNum),
								Notes = TryGetString(currentRow.GetCell(colIndexNotes), workSheetName, cellHeaderNotes.StringCellValue, rowNum)
							};

							if (plans.Any(x => x.TestCaseId.Equals(param.TestCaseId)))
								result.Add(param);
						}
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read Sheet {SheetConstant.Task} is Fail : {CheckErrorReadExcel(ex.Message)}");
			}
		}

		private static List<TestplanModels> ReadTestPlans(IWorkbook wb)
		{
			List<TestplanModels> result = [];
			try
			{
				string workSheetName = SheetConstant.TestPlan;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexTestCaseId = 0;
				int colIndexModuleName = colIndexTestCaseId + 1;
				int colIndexTestCase = colIndexModuleName + 1;
				int colIndexFromTestCaseId = colIndexTestCase + 1;
				int colIndexStatus = colIndexFromTestCaseId + 1;
				int colIndexDateTest = colIndexStatus + 1;
				int colIndexRemarks = colIndexDateTest + 1;
				int colIndexScreenCapture = colIndexRemarks + 1;
				int colIndexTestData = colIndexScreenCapture + 1;
				int colIndexReqNum = colIndexTestData + 1;
				int colIndexUserLogin = colIndexReqNum + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderTestCaseId = firstRow.GetCell(colIndexTestCaseId);
				ICell cellHeaderModuleName = firstRow.GetCell(colIndexModuleName);
				ICell cellHeaderTestCase = firstRow.GetCell(colIndexTestCase);
				ICell cellHeaderFromTestCaseId = firstRow.GetCell(colIndexFromTestCaseId);
				ICell cellHeaderStatus = firstRow.GetCell(colIndexStatus);
				ICell cellHeaderDateTest = firstRow.GetCell(colIndexDateTest);
				ICell cellHeaderRemarks = firstRow.GetCell(colIndexRemarks);
				ICell cellHeaderScreenCapture = firstRow.GetCell(colIndexScreenCapture);
				ICell cellHeaderTestData = firstRow.GetCell(colIndexTestData);
				ICell cellHeaderReqNum = firstRow.GetCell(colIndexReqNum);
				ICell cellHeaderUserLogin = firstRow.GetCell(colIndexUserLogin);

				ValidateCellHeaderName(cellHeaderTestCaseId, workSheetName, colIndexTestCaseId + 1, "Test Case Id");
				ValidateCellHeaderName(cellHeaderModuleName, workSheetName, colIndexModuleName + 1, "Module Name");
				ValidateCellHeaderName(cellHeaderTestCase, workSheetName, colIndexTestCase + 1, "Test Case");
				ValidateCellHeaderName(cellHeaderStatus, workSheetName, colIndexStatus + 1, "Status");
				ValidateCellHeaderName(cellHeaderDateTest, workSheetName, colIndexDateTest + 1, "Date Test");
				ValidateCellHeaderName(cellHeaderRemarks, workSheetName, colIndexRemarks + 1, "Remarks");
				ValidateCellHeaderName(cellHeaderScreenCapture, workSheetName, colIndexScreenCapture + 1, "Screen Capture");
				ValidateCellHeaderName(cellHeaderTestData, workSheetName, colIndexTestData + 1, "Test Data");
				ValidateCellHeaderName(cellHeaderUserLogin, workSheetName, colIndexUserLogin + 1, "User Login");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (IsRowEmpty(currentRow, colIndexUserLogin))
					{
						break;
					}

					#region Assign Param
					string id = TryGetString(currentRow.GetCell(colIndexTestCaseId), workSheetName, cellHeaderTestCaseId.StringCellValue, rowNum);
					var status = TryGetString(currentRow.GetCell(colIndexStatus), workSheetName, cellHeaderStatus.StringCellValue, rowNum);
					string moduleName = TryGetString(currentRow.GetCell(colIndexModuleName), workSheetName, cellHeaderModuleName.StringCellValue, rowNum);
					string testCase = TryGetString(currentRow.GetCell(colIndexTestCase), workSheetName, cellHeaderTestCase.StringCellValue, rowNum);
					string fromTesctCaseId = TryGetString(currentRow.GetCell(colIndexFromTestCaseId), workSheetName, cellHeaderFromTestCaseId.StringCellValue, rowNum);

					if (!string.IsNullOrEmpty(id))
					{
						if (!string.IsNullOrEmpty(moduleName) || !string.IsNullOrEmpty(testCase))
						{
							if (!string.IsNullOrEmpty(testCase))
							{
								TestplanModels testPlan = new()
								{
									Row = rowNum,
									TestCaseId = int.Parse(id),
									FromTestCaseId = !string.IsNullOrEmpty(fromTesctCaseId) ? int.Parse(fromTesctCaseId) : 0,
									ModuleName = moduleName,
									TestCase = testCase,
									Status = status,
									RequestNumber = TryGetString(currentRow.GetCell(colIndexReqNum), workSheetName, cellHeaderReqNum.StringCellValue, rowNum),
									UserLogin = TryGetString(currentRow.GetCell(colIndexUserLogin), workSheetName, cellHeaderUserLogin.StringCellValue, rowNum)
								};
								result.Add(testPlan);
							}
						}
					}
					#endregion
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Read Sheet {SheetConstant.TestPlan} is Fail : {CheckErrorReadExcel(ex.Message)}");
			}
		}

		private static List<UserModels> ReadUsers(IWorkbook wb)
		{
			List<UserModels> result = [];
			try
			{
				string workSheetName = SheetConstant.User;
				ISheet sheet = wb.GetSheet(workSheetName);

				#region Validate Column
				int colIndexId = 0;
				int colIndexUsername = colIndexId + 1;
				int colIndexPassword = colIndexUsername + 1;
				int colIndexRole = colIndexPassword + 1;

				IRow firstRow = sheet.GetRow(0);
				ICell cellHeaderId = firstRow.GetCell(colIndexId);
				ICell cellHeaderUsername = firstRow.GetCell(colIndexUsername);
				ICell cellHeaderPassword = firstRow.GetCell(colIndexPassword);
				ICell cellHeaderRole = firstRow.GetCell(colIndexRole);

				ValidateCellHeaderName(cellHeaderId, workSheetName, colIndexId + 1, "Id");
				ValidateCellHeaderName(cellHeaderUsername, workSheetName, colIndexUsername + 1, "Username");
				ValidateCellHeaderName(cellHeaderPassword, workSheetName, colIndexPassword + 1, "Password");
				ValidateCellHeaderName(cellHeaderRole, workSheetName, colIndexRole + 1, "Role");
				#endregion

				int rowNum = 0;

				while (true)
				{
					rowNum++;
					IRow? currentRow = sheet.GetRow(rowNum);
					if (currentRow == null)
					{
						break;
					}
					if (IsRowEmpty(currentRow, colIndexRole))
					{
						break;
					}

					UserModels user = new();
					#region Assign Param
					user.Id = int.Parse(TryGetString(currentRow.GetCell(colIndexId), workSheetName, cellHeaderId.StringCellValue, rowNum));
					user.Username = TryGetString(currentRow.GetCell(colIndexUsername), workSheetName, cellHeaderUsername.StringCellValue, rowNum);
					user.Password = TryGetString(currentRow.GetCell(colIndexPassword), workSheetName, cellHeaderPassword.StringCellValue, rowNum);
					user.Role = TryGetString(currentRow.GetCell(colIndexRole), workSheetName, cellHeaderRole.StringCellValue, rowNum);
					#endregion
					result.Add(user);
				}
				return result;
			}
			catch (Exception ex)
			{
				throw new Exception($"Fail Read Sheet {SheetConstant.User} : {ex.Message}");
			}
		}
		#endregion

		#region Write Excel		
		public static async void WriteAutomationResult(AppConfig cfg, TestplanModels param)
		{
			try
			{
				string dateNow = DateTime.Now.ToString("yyyy/MM/dd");
				string dateTimeNow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

				using (var fs = new FileStream(cfg.ExcelPath, FileMode.Open, FileAccess.Read))
				{
					var wb = new XSSFWorkbook(fs);
					ISheet sheet = wb.GetSheet("TestPlan");

					IRow row = sheet.GetRow(param.Row);
					var cell4 = row.CreateCell(4);
					cell4.SetCellValue(param.Status);
					cell4.CellStyle.Alignment = HorizontalAlignment.Center;
					var cell5 = row.CreateCell(5);
					cell5.SetCellValue(dateTimeNow);
					cell5.CellStyle.Alignment = HorizontalAlignment.Center;
					var cell6 = row.CreateCell(6);
					cell6.SetCellValue(param.Remarks);
					cell6.CellStyle.Alignment = HorizontalAlignment.Left;

					string filePath = string.Empty;
					try
					{
						filePath = await ScreenCapture(cfg, param);
					}
					catch (Exception ex)
					{
						Console.WriteLine($"Screen Capture failed: {ex.Message}");
					}

					if (!string.IsNullOrEmpty(filePath))
					{
						var cellScreenCapture = row.CreateCell(7);
						ICreationHelper creationHelper = wb.GetCreationHelper();
						IHyperlink hyperlink = creationHelper.CreateHyperlink(HyperlinkType.File);
						hyperlink.Address = filePath;
						cellScreenCapture.Hyperlink = hyperlink;
						cellScreenCapture.CellStyle.Alignment = HorizontalAlignment.Center;
						cellScreenCapture.SetCellValue("Open File");
					}

					if (!string.IsNullOrEmpty(param.TestData))
					{
						var testData = row.CreateCell(8);
						ICreationHelper creationHelper = wb.GetCreationHelper();
						IHyperlink hyperlink = creationHelper.CreateHyperlink(HyperlinkType.Document);
						hyperlink.Address = param.TestData;
						testData.Hyperlink = hyperlink;
						testData.CellStyle.Alignment = HorizontalAlignment.Center;
						testData.SetCellValue("Go To Data");
					}

					var cell9 = row.CreateCell(9);
					cell9.SetCellValue(param.RequestNumber);
					cell9.CellStyle.Alignment = HorizontalAlignment.Center;

					using (var outputStream = new FileStream(cfg.ExcelPath, FileMode.Create, FileAccess.Write))
					{
						wb.Write(outputStream);
						outputStream.Close();
					}
					fs.Close();
					wb.Close();
					Thread.Sleep(1500);
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Fail WriteAutomationResult : {ex.Message}");
			}
		}
		public static async void WriteApproval<T>(AppConfig cfg, T param, string moduleName) where T : class
		{
			try
			{
				string dateTimeNow = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

				using (var fs = new FileStream(cfg.ExcelPath, FileMode.Open, FileAccess.Read))
				{
					var wb = new XSSFWorkbook(fs);
					ISheet sheet;
					int resultColumnIndex, captureColumnIndex, rowIndex;
					string testCase, status;

					if (param is TaskModels taskParam)
					{
						sheet = wb.GetSheet(SheetConstant.Task);
						resultColumnIndex = 5;
						captureColumnIndex = 6;
						TaskModels newParam = param as TaskModels ?? new();
						rowIndex = newParam.Row;
						testCase = $"{TestCaseConstant.Approval} - {newParam.Action}";
						status = newParam.Result == StatusConstant.Success ? $"{StatusConstant.Success} - {newParam.Action} with Document Number : {newParam.DataNumber}" : newParam.Result;
					}
					else
					{
						throw new InvalidOperationException("Unsupported parameter type.");
					}

					var row = sheet.GetRow(rowIndex);
					ICell cellResult = row.CreateCell(resultColumnIndex);
					cellResult.CellStyle.Alignment = HorizontalAlignment.Left;
					cellResult.SetCellValue(status);

					TestplanModels plan = new()
					{
						TestCase = testCase,
						ModuleName = moduleName,
						Status = status.Contains("Error") ? "Fail" : "Success"
					};

					string filePath = string.Empty;
					try
					{
						filePath = await ScreenCapture(cfg, plan);
					}
					catch (Exception ex)
					{
						Console.WriteLine($"Screen Capture failed: {ex.Message}");
					}

					if (!string.IsNullOrEmpty(filePath))
					{
						var cellScreenCapture = row.CreateCell(captureColumnIndex);
						var creationHelper = wb.GetCreationHelper();
						var hyperlink = creationHelper.CreateHyperlink(HyperlinkType.File);
						hyperlink.Address = filePath;
						cellScreenCapture.Hyperlink = hyperlink;
						cellScreenCapture.CellStyle.Alignment = HorizontalAlignment.Center;
						cellScreenCapture.SetCellValue("Open File");
					}

					using (var outputStream = new FileStream(cfg.ExcelPath, FileMode.Create, FileAccess.Write))
					{
						wb.Write(outputStream);
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Fail WriteApproval: {ex.Message}");
			}
		}
		#endregion

		#region Private Function
		[DllImport("user32.dll")]
		private static extern int GetSystemMetrics(int nIndex);

		private const int SM_CXSCREEN = 0;
		private const int SM_CYSCREEN = 1;

		private static int IntegerParse(string param)
		{
			return int.Parse(!string.IsNullOrEmpty(param) ? param : "0");
		}
		private static async Task<string> ScreenCapture(AppConfig cfg, TestplanModels param)
		{
			try
			{
				string dateNow = DateTime.Now.ToString("yyyy-MM-dd");
				string dateTime = DateTime.Now.ToString("HHmmss");
				var filePath = Path.Combine(dateNow, $"{dateTime}_{param.TestCase}_{param.ModuleName}_{param.Status}.png");
				string fullFilePath = Path.Combine(cfg.ScreenCapturePath, filePath);
				DirectoryCheck(fullFilePath);

				// Capture the screen
				int screenWidth = GetSystemMetrics(SM_CXSCREEN);
				int screenHeight = GetSystemMetrics(SM_CYSCREEN);

				using var bitmap = new Bitmap(screenWidth, screenHeight);
				using var graphics = Graphics.FromImage(bitmap);
				graphics.CopyFromScreen(0, 0, 0, 0, bitmap.Size);
				bitmap.Save(fullFilePath, System.Drawing.Imaging.ImageFormat.Png);

				return $"ScreenCapture\\{filePath}";
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Failed to get screen capture {param.ModuleName} - {param.TestCase}: {ex.Message}");
				return "";
			}
		}
		private static void DirectoryCheck(string filePath)
		{
			var directory = Path.GetDirectoryName(filePath);
			if (!Directory.Exists(directory))
				Directory.CreateDirectory(directory!);
		}
		private static string CheckErrorReadExcel(string error)
		{
			return error.Contains("Input string was not in a correct format or must be in numeric format.") ? $"{error} Test Case Id or Sequence cannot be empty." : error;
		}
		private static bool ValidateCellHeaderName(ICell cellHeader, string worksheetName, int cellHeaderDisplayIndex, string validHeaderNames)
		{
			string cellHeaderValue = cellHeader.StringCellValue.Trim();

			if (cellHeaderValue.Equals(validHeaderNames, StringComparison.OrdinalIgnoreCase))
			{
				return true;
			}
			else
			{
				Console.WriteLine($"{worksheetName} - Column {cellHeaderDisplayIndex.ToString()} must be {validHeaderNames}");
				return false;
			}
		}

		private static string TryGetString(ICell? cell, string worksheetName, string headerName, int rowNum)
		{
			string result = string.Empty;
			headerName = headerName.Trim();
			if (cell == null)
			{
				//Console.WriteLine($"{worksheetName} - Column {headerName} - Row {rowNum} is empty.");
			}
			else
			{
				return TryGetStringAllowNull(cell, worksheetName, headerName, rowNum) ?? string.Empty;
			}
			return result;
		}

		private static string? TryGetStringAllowNull(ICell? cell, string worksheetName, string headerName, int rowNum)
		{
			string result = string.Empty;
			headerName = headerName.Trim();
			if (cell == null)
			{
				return null;
			}
			else
			{
				switch (cell.CellType)
				{
					case CellType.Numeric:
						result = cell.NumericCellValue.ToString();
						break;
					case CellType.String:
						result = cell.StringCellValue;
						break;
					case CellType.Boolean:
						result = cell.BooleanCellValue.ToString();
						break;
					case CellType.Unknown:
					case CellType.Formula:
					case CellType.Blank:
						result = cell.ToString() ?? string.Empty;
						break;
					case CellType.Error:
					default:
						Console.WriteLine($"{worksheetName} - An error occured when trying to get the value of Column {headerName} - Row {rowNum} is empty.");
						break;
				}
			}
			return result;
		}

		private static bool IsRowEmpty(IRow row, int maxColumnIndex)
		{
			bool allColumnsEmpty = true;
			for (int colIndex = 0; colIndex <= maxColumnIndex; colIndex++)
			{
				if (row.GetCell(colIndex) != null && row.GetCell(colIndex).CellType != CellType.Blank)
				{
					allColumnsEmpty = false;
					break;
				}
			}

			return allColumnsEmpty;
		}

		private static bool TryGetBoolean(ICell? cell, string worksheetName, string headerName, int rowNum)
		{
			bool result = false;
			headerName = headerName.Trim();
			if (cell == null)
			{
				//Console.WriteLine($"{worksheetName} - Column {headerName} - Row {rowNum} is empty.");
			}
			else
			{
				return TryGetBooleanAllowNull(cell, worksheetName, headerName, rowNum) ?? false;
			}
			return result;
		}
		private static bool? TryGetBooleanAllowNull(ICell? cell, string worksheetName, string headerName, int rowNum)
		{
			bool? result = null;
			headerName = headerName.Trim();
			if (cell != null)
			{
				try
				{
					switch (cell.CellType)
					{
						case CellType.Numeric:
							result = Convert.ToBoolean(cell.NumericCellValue);
							break;
						case CellType.String:
							result = Convert.ToBoolean(cell.StringCellValue);
							break;
						case CellType.Boolean:
							result = cell.BooleanCellValue;
							break;
						case CellType.Unknown:
						case CellType.Formula:
						case CellType.Blank:
							break;
						case CellType.Error:
						default:
							Console.WriteLine($"{worksheetName} - An error occured when trying to get the value of Column {headerName} - Row {rowNum} is empty.");
							break;
					}
				}
				catch
				{
					Console.WriteLine($"{worksheetName} - An error occured when trying to get the value of Column {headerName} - Row {rowNum} is empty.");
				}
			}
			return result;
		}

		#endregion
	}
}
