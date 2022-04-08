using IntegrationAutomation.Excel;
using IntegrationAutomation.Selenium;

namespace IntegrationAutomation
{
    public class IntegrationAutomationOperationFactory
    {
        public static IntegrationAutomationOperation CreateInstance(AppSettingClass appSettingClass)
        {
            OperationInterface operationInterface = new SeleniumHelper(new TxtLogWriter(), appSettingClass.ChromeDriverFilePath);
            OperationInterface excelInterface = new ExcelHelper(new TxtLogWriter());
            OperationInterface systemInterface = new SystemHelper(new TxtLogWriter());
            return new IntegrationAutomationOperation(operationInterface, excelInterface, systemInterface);
        }
    }
}
