using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegrationAutomation.Selenium
{
    public class SeleniumHelper : OperationInterface
    {
        WebDriver driver;
        public void PerforAction(ActionRowEntity action)
        {
            switch (action.CommandName)
            {
                case "GoToUrl":
                    var browserName = action.CommandComponent;
                    switch (browserName)
                    {
                        case "Chrome":
                            driver = new ChromeDriver(@"C:\D\chromedriver_win32");
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
                    driver.FindElement(By.Id(action.CommandComponent)).SendKeys(action.CommandValue);
                    break;
                case "Click":
                    driver.FindElement(By.Id(action.CommandComponent)).Click();
                    break;
                default:
                    throw new ArgumentException("Command Type is not implemented");
            }

            WaitTime(action.WaitTime);
        }

        private void WaitTime(string time)
        {
            if(!string.IsNullOrEmpty(time))
            {
                if(Int32.TryParse(time, out int minutes))
                {
                    Thread.Sleep(minutes);
                }
                else
                {
                    throw new ArgumentException("Wait time is not an integer");
                }
            }
        }
    }
}
