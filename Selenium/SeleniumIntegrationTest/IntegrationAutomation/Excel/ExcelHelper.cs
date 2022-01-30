using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegrationAutomation.Excel
{
    public class ExcelHelper : OperationInterface
    {
        ILog _log { get; set; }
        public ExcelHelper(ILog log)
        {
            _log = log;
        }

        public void PerforAction(ActionRowEntity action)
        {
            _log.Log("Before Action: " + action.ToString());
            try
            {
                switch (action.CommandName)
                {
                    case "Compare File":
                        var browserName = action.CommandComponent;

                        break;
                    default:
                        throw new ArgumentException("Command Type is not implemented");
                }
            }
            catch (Exception ex)
            {
                _log.Log("Exception:" + ex.ToString());
                throw new Exception(ex.ToString());
            }

            WaitTime(action.WaitTime);
        }

        public void WaitTime(string time)
        {
            if (!string.IsNullOrEmpty(time))
            {
                if (Int32.TryParse(time, out int seconds))
                {
                    _log.Log("Wait Seconds:" + seconds.ToString());
                    //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(seconds);
                    Thread.Sleep(seconds * 1000);
                }
                else
                {
                    throw new ArgumentException("Wait time is not an integer");
                }
            }
        }
    }
}
