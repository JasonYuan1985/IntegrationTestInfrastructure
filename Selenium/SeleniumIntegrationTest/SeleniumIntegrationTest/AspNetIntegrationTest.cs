using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Drawing;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;

namespace SeleniumIntegrationTest
{
    public class AspNetIntegrationTest
    {
        public void TestUpload()
        {
            new DriverManager().SetUpDriver(new ChromeConfig());

            var driver = new ChromeDriver(@"C:\D\chromedriver_win32");

            var url = "www.google.com";
            driver.Navigate().GoToUrl(url);

            driver.Manage().Window.Size = new Size(1440, 900);

            var title = driver.Title;

            if (title == "用户登陆")
            {
                var txtUserName = driver.FindElement(By.Id("txtUser"));

                txtUserName.SendKeys("test");

                var txtPassword = driver.FindElement(By.Id("txtPassword"));

                txtPassword.SendKeys("test");

                var btnLogin = driver.FindElement(By.Id("btnLogin1"));

                btnLogin.Click();
            }

            System.Threading.Thread.Sleep(5000);

            var title_Home = driver.Title;

            var tabSystemSetting = driver.FindElement(By.XPath("//*[@id=\"ctl00_ucHeader1_lnkbtnSystemCenter\"]"));

            tabSystemSetting.Click();

            System.Threading.Thread.Sleep(2000);

            var ucEmployeeCost = driver.FindElement(By.XPath("//*[@id=\"divMenuSystem\"]/table/tbody/tr[36]/td[2]/a"));

            ucEmployeeCost.Click();

            System.Threading.Thread.Sleep(2000);

            var ucFileUpload = driver.FindElement(By.XPath("//*[@id=\"ctl00_ctl00_cphContent_cphContent_FileUpload1\"]"));

            ucFileUpload.SendKeys(@"C:\Users\Jason_Yuan\Downloads\Cost.xlsx");

            System.Threading.Thread.Sleep(2000);

            var btnUpload = driver.FindElement(By.XPath("//*[@id=\"ctl00_ctl00_cphContent_cphContent_btnUpload\"]"));

            btnUpload.Click();

            driver.Quit();

            Console.WriteLine(title + "," + title_Home);
        }
    }
}
