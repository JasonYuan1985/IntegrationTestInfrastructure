// See https://aka.ms/new-console-template for more information
using IntegrationAutomation;
using IntegrationAutomation.Selenium;
using System.Data;

ExcelReader excelReader = new ExcelReader();
string activeSheetName;
string excelFileName = @"C:\Users\Jason_Yuan\Downloads\integration test import template.xlsx";
DataSet dataSet = excelReader.ExcelToDataSet(excelFileName, "", 0, out activeSheetName);
DataTable dataTable = dataSet.Tables[0];
//convert to ActionRowEntity
OperationInterface operationInterface = new SeleniumHelper();
IntegrationAutomationOperation operation = new IntegrationAutomationOperation(operationInterface);
if(operation.CheckTemplate(dataTable))
{
    var entities = operation.GetActionRowEntities(dataTable);
    operation.PerformActions(entities);
}
else
{
    Console.WriteLine("Template Error");
}
