using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Threading;

namespace Test_Lab_1
{
    class Program
    {
        static void Main(string[] args)
        {
            IWebDriver driver = new ChromeDriver();
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            driver.Navigate().GoToUrl("http://Capcrmdev0\*****:****@**.**.**.179:5555/****/main.aspx#433885818");
            driver.SwitchTo().Frame("InlineDialog1_Iframe");
            driver.FindElement(By.XPath("//*[@id='butBegin']")).Click();
            driver.SwitchTo().Frame(1);
            driver.Manage().Timeouts().PageLoad = (TimeSpan.FromSeconds(5));
            driver.FindElement(By.XPath("//*[@id='butBegin']")).Click();
            driver.Manage().Window.Maximize();
            driver.SwitchTo().ParentFrame();
            driver.FindElement(By.Id("HomeTabLink")).Click();
            IWebElement e = driver.FindElement(By.XPath("//a[@id='SFA']/span/span/img"));
            Thread.Sleep(1000);
            e.Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Id("tempId_636190677158102561")).Click();
            Thread.Sleep(1000);
            var navbar = driver.FindElement(By.Id("navBarOverlay"));
            js.ExecuteScript("arguments[0].setAttribute('style', 'display: none')", navbar);
            driver.FindElement(By.XPath("//*[@id='im_networkpartner|NoRelationship|HomePageGrid|Mscrm.HomepageGrid.im_networkpartner.NewRecord']/span/a")).Click();
            Thread.Sleep(2000);
            driver.SwitchTo().Frame("contentIFrame1");
            js.ExecuteScript("Xrm.Page.getAttribute('im_name').setValue('Selenium Test');");
            js.ExecuteScript("Xrm.Page.getAttribute('im_dunsnumber').setValue('1223323');Xrm.Page.getAttribute('im_networkpartner').setValue('NWP_1234');Xrm.Page.getAttribute('im_networkpartnershortname').setValue('NWP_12')");
            js.ExecuteScript("var lookupValue = new Array();lookupValue[0] = new Object();lookupValue[0].id = '{B25E829B-26BD-E611-80F3-000D3AA17340}';lookupValue[0].name = 'Argentina';lookupValue[0].entityType = 'im_country';Xrm.Page.getAttribute('im_country').setValue(lookupValue)");
            driver.SwitchTo().ParentFrame();
            Thread.Sleep(1000);
            var save= driver.FindElement(By.XPath("//*[@id='im_networkpartner|NoRelationship|Form|Mscrm.Form.im_networkpartner.Save']/span/a/span"));
            save.Click();
        }
    }
}
