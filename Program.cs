﻿using LatihanSelenium.Models;
using LatihanSelenium.Utilities;

using Newtonsoft.Json;
using System.Text;

#region App Setting
var basePath = "E:\\My Data\\Project\\selenium\\dotnet\\LatihanSelenium";

string jsonString = File.ReadAllText($"{basePath}\\appsettings.json", Encoding.Default);
AppConfig cfg = JsonConvert.DeserializeObject<AppConfig>(jsonString)!;
cfg.ExcelPath = $"{basePath}{cfg.ExcelPath}";
cfg.ScreenCapturePath = $"{basePath}{cfg.ScreenCapturePath}";
GlobalConfig.Config = cfg;
#endregion

Console.WriteLine($"Running Unit Test");
Console.WriteLine("------------------------------------------------------");

UnitTestHandler.RunUnitTest();

Console.WriteLine($"End of Unit Test Sequence");
Console.WriteLine("------------------------------------------------------");