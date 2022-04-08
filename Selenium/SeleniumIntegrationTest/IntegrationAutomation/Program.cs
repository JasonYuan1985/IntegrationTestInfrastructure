// See https://aka.ms/new-console-template for more information
using IntegrationAutomation;
using Microsoft.Extensions.Configuration;
using System.Data;

// Build a config object, using env vars and JSON providers.
IConfiguration config = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json")
    .AddEnvironmentVariables()
    .Build();

// Get values from the config given their key and their target type.
var appSetting = config.GetSection("ApplicationSettings").Get<AppSettingClass>();

ExcelReader excelReader = new ExcelReader();
string activeSheetName;
string excelFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, appSetting.ExcelFileName);
DataSet dataSet = excelReader.ExcelToDataSet(excelFileName, "", 0, out activeSheetName);
DataTable dataTable = dataSet.Tables[0];
//convert to ActionRowEntity
IntegrationAutomationOperation operation = IntegrationAutomationOperationFactory.CreateInstance(appSetting);
if (operation.CheckTemplate(dataTable))
{
    var entities = operation.GetActionRowEntities(dataTable);
    operation.PerformActions(entities);
}
else
{
    Console.WriteLine("Template Error");
}
operation.CloseAutomation();
