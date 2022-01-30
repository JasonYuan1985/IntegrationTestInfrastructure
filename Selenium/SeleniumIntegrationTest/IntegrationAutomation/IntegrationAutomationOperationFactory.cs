using IntegrationAutomation.Excel;
using IntegrationAutomation.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegrationAutomation
{
    public class IntegrationAutomationOperationFactory
    {
        public static IntegrationAutomationOperation CreateInstance()
        {
            OperationInterface operationInterface = new SeleniumHelper(new TxtLogWriter());
            OperationInterface excelInterface = new ExcelHelper(new TxtLogWriter());
            return new IntegrationAutomationOperation(operationInterface, excelInterface);
        }
    }
}
