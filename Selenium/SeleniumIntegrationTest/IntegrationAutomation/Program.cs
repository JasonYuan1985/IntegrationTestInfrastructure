// See https://aka.ms/new-console-template for more information
using IntegrationAutomation;
using System.Data;

ExcelReader excelReader = new ExcelReader();
string activeSheetName;
string excelFileName = @"C:\Users\Jason_Yuan\Downloads\integration test import template.xlsx";
DataSet dataSet = excelReader.ExcelToDataSet(excelFileName, "", 0, out activeSheetName);
DataTable dataTable = dataSet.Tables[0];
//convert to ActionRowEntity
IntegrationAutomationOperation operation = IntegrationAutomationOperationFactory.CreateInstance();
if(operation.CheckTemplate(dataTable))
{
    var entities = operation.GetActionRowEntities(dataTable);
    operation.PerformActions(entities);
}
else
{
    Console.WriteLine("Template Error");
}
