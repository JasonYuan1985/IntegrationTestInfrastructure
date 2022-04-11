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
        string _chromeDriverPath { get; set; }
        public SeleniumHelper(ILog log, string chromeDriverPath)
        {
            _log = log;
            _chromeDriverPath = chromeDriverPath;
        }

        public void PerforAction(ActionRowEntity action)
        {
            _log.Log("Before Action: " + action.ToString());
            bool needWaitTime = true;
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
                                driver = new ChromeDriver(_chromeDriverPath, options);
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
                    case "GetValue":
                        bool result = IsTargetValueEqualToElement(driver, by, action.CommandValue, action.WaitTime);
                        needWaitTime = false;
                        if (!result) throw new ArgumentException("Not found value");
                        break;
                    default:
                        throw new ArgumentException("Command Name is not implemented");
                }
            }
            catch (Exception ex)
            {
                _log.Log("Exception:" + ex.ToString());
                throw new Exception(action.ToString() + "\r\n" + ex.ToString());
            }

            if (needWaitTime) WaitTime(action.WaitTime);
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
            if (driver != null)
            {
                driver.Quit();
            }
        }

        private bool IsTargetValueEqualToElement(WebDriver driver, By byPath, string value, string maxWaitTime)
        {
            try
            {
                //3秒钟做获取alert的内容
                var labelText = GetText(driver.FindElement(byPath));
                int eachWaitTime = 1000;
                int totalWaitTime = 0;
                int maxWaitTimeInteger = string.IsNullOrEmpty(maxWaitTime) ? 0 : Convert.ToInt32(maxWaitTime);
                //每秒获取一次
                while (labelText != value && totalWaitTime < maxWaitTimeInteger * 1000)
                {
                    Thread.Sleep(eachWaitTime);
                    totalWaitTime += eachWaitTime;
                    labelText = GetText(driver.FindElement(byPath));
                }

                return labelText == value;
            }
            catch (Exception ex)
            {
                _log.Log("查找文件报错:" + ex.ToString());
                return false;
            }
        }

        private string GetText(IWebElement webElement)
        {
            _log.Log("Text:" + webElement.ToString());
            return webElement.Text;
        }
    }
}
