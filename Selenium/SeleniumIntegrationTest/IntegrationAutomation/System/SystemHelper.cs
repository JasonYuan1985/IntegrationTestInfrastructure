using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Diagnostics;
using System.Drawing;

namespace IntegrationAutomation.Selenium
{
    public class SystemHelper : OperationInterface
    {
        ILog _log { get; set; }
        public SystemHelper(ILog log)
        {
            _log = log;
        }

        public void PerforAction(ActionRowEntity action)
        {
            _log.Log("Before Action: " + action.ToString());
            var actionValue = action.CommandValue;
            var actionCommand = action.CommandName;
            var actionType = action.CommandType;
            try
            {
                switch (actionCommand)
                {
                    case "Set Clipboard":
                        TextCopy.ClipboardService.SetText(actionValue);
                        break;
                    case "SendKeys":
                        Process.Start("SendKeysTool.exe", actionValue);
                        break;
                    default:
                        throw new ArgumentException("Command Type is not implemented");
                }
            }
            catch (Exception ex)
            {
                _log.Log("Exception:" + ex.ToString());
                throw new Exception(action.ToString() + "\r\n" + ex.ToString());
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

        public void Close()
        {
            
        }
    }
}
