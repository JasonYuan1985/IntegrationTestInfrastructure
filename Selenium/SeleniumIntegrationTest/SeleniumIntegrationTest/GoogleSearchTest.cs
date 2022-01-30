using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Drawing;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;

namespace SeleniumIntegrationTest
{
    public class GoogleSearchTest
    {
        public void TestSearch()
        {
            new DriverManager().SetUpDriver(new ChromeConfig());

            var driver = new ChromeDriver(@"C:\D\chromedriver_win32");

            var url = "https://www.google.com";
            driver.Navigate().GoToUrl(url);

            driver.Manage().Window.Size = new Size(1440, 900);

            var searchBox = driver.FindElement(By.Name("q"));
            var searchButton = driver.FindElement(By.Name("btnK"));

            searchBox.SendKeys("Selenium");

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);

            searchButton.Click();


            var title = driver.Title;

            driver.Quit();

            Console.WriteLine(title);
        }
    }
}
