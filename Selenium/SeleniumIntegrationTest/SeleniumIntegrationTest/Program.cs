// See https://aka.ms/new-console-template for more information


// Import the dependencies:
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Drawing;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;

new DriverManager().SetUpDriver(new ChromeConfig());

//Chrome version: 96.0.4664.110
//driver version: 96.0.4664.18
var driver = new ChromeDriver(@"C:\D\chromedriver_win32");
//"https://www.google.com"
var url = "";
driver.Navigate().GoToUrl(url);


Size d = new Size(1440, 900);
//Resize the current window to the given dimension
driver.Manage().Window.Size = d;

var title = driver.Title;

if(title == "Sign On")
{
    var txtUserName = driver.FindElement(By.Id("username"));

    txtUserName.SendKeys("test");

    var txtPassword = driver.FindElement(By.Id("password"));

    txtPassword.SendKeys("test");

    var btnOK = driver.FindElement(By.Id("loginButton"));

    btnOK.Click();
}

System.Threading.Thread.Sleep(5000);

var title_Home = driver.Title;

var tabMaster = driver.FindElement(By.XPath("//span[contains(.,'OOS Master')]"));

tabMaster.Click();

System.Threading.Thread.Sleep(2000);

var tabDCName = driver.FindElement(By.XPath("//span[contains(.,'DC Name')]"));

tabDCName.Click();

System.Threading.Thread.Sleep(2000);

var btnExportDCName = driver.FindElement(By.XPath("//*[@id=\"app\"]/div/div[2]/section/div/div[1]/div/button[2]"));

btnExportDCName.Click();

System.Threading.Thread.Sleep(2000);

driver.Quit();

Console.WriteLine(title + "," + title_Home);
