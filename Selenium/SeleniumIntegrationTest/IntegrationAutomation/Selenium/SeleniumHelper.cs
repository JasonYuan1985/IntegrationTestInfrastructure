using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System.Drawing;

namespace IntegrationAutomation.Selenium
{
    public class SeleniumHelper : OperationInterface
    {
        WebDriver driver;
        ILog _log { get; set; }
        public SeleniumHelper(ILog log)
        {
            _log = log;
        }

        public void PerforAction(ActionRowEntity action)
        {
            _log.Log("Before Action: " + action.ToString());

            try
            {

                By by = null;
                if (!string.IsNullOrEmpty(action.CommandType))
                {
                    switch (action.CommandType.ToLower())
                    {
                        case "id":
                            by = By.Id(action.CommandComponent);
                            break;
                        case "xpath":
                            by = By.XPath(action.CommandComponent);
                            break;
                    }
                }

                switch (action.CommandName)
                {
                    case "GoToUrl":
                        var browserName = action.CommandComponent;
                        switch (browserName)
                        {
                            case "Chrome":
                                ChromeOptions options = new ChromeOptions();
                                options.AcceptInsecureCertificates = true;
                                driver = new ChromeDriver(@"C:\D\chromedriver_win32", options);
                                break;
                            default:
                                throw new ArgumentException("Not inplement browser name");
                        }
                        driver.Navigate().GoToUrl(action.CommandValue);
                        break;
                    case "SetSize":
                        List<string> sizeLists = action.CommandValue.Split(",").Select(s => s.Trim()).ToList();
                        driver.Manage().Window.Size = new Size(Convert.ToInt32(sizeLists.FirstOrDefault())
                            , Convert.ToInt32(sizeLists.LastOrDefault()));
                        break;
                    case "SetValue":
                        driver.FindElement(by).SendKeys(action.CommandValue);
                        break;
                    case "Click":
                        driver.FindElement(by).Click();
                        break;
                    case "SendKeys":
                        Actions keyAction = new Actions(driver);
                        keyAction.SendKeys(OpenQA.Selenium.Keys.Enter).Build().Perform();
                        break;
                    default:
                        throw new ArgumentException("Command Name is not implemented");
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
