using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using System.Threading;

namespace QCTTestCase
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            IWebDriver driver = new ChromeDriver();
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            driver.Navigate().GoToUrl("http://****\arindhsa:Zurich_2016!@13.*****:5555/STORMDev/main.aspx#433885818");
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
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//*[@id='im_networkpartner|NoRelationship|HomePageGrid|Mscrm.HomepageGrid.im_networkpartner.NewRecord']/span/a")).Click();
            Thread.Sleep(2000);
            driver.SwitchTo().Frame("contentIFrame1");
            js.ExecuteScript("Xrm.Page.getAttribute('im_name').setValue('Selenium Test');");
            js.ExecuteScript("Xrm.Page.getAttribute('im_dunsnumber').setValue('1223323');Xrm.Page.getAttribute('im_networkpartner').setValue('NWP_1234');Xrm.Page.getAttribute('im_networkpartnershortname').setValue('NWP_12')");
            js.ExecuteScript("var lookupValue = new Array();lookupValue[0] = new Object();lookupValue[0].id = '{B25E829B-26BD-E611-80F3-000D3AA17340}';lookupValue[0].name = 'Argentina';lookupValue[0].entityType = 'im_country';Xrm.Page.getAttribute('im_country').setValue(lookupValue)");
            driver.SwitchTo().ParentFrame();
            Thread.Sleep(1000);
            var save = driver.FindElement(By.XPath("//*[@id='im_networkpartner|NoRelationship|Form|Mscrm.Form.im_networkpartner.Save']/span/a/span"));
            save.Click();
        }

        [TestMethod]
        public void TestMethod2()
        {
            IWebDriver driver = new ChromeDriver();
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            driver.Navigate().GoToUrl("http://******\arindhsa:Zurich_2016!@********:5555/STORMDev/main.aspx#433885818");
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
            driver.FindElement(By.Id("navProjects")).Click();
            Thread.Sleep(1000);
            var navbar = driver.FindElement(By.Id("navBarOverlay"));
            js.ExecuteScript("arguments[0].setAttribute('style', 'display: none')", navbar);
            Thread.Sleep(2000);
            var newButton = driver.FindElement(By.XPath("//*[@id='im_project|NoRelationship|HomePageGrid|Mscrm.HomepageGrid.im_project.NewRecord']/span/a"));
            Thread.Sleep(2000);
            newButton.Click();
            Thread.Sleep(2000);
            driver.SwitchTo().Frame("contentIFrame1");
            js.ExecuteScript("Xrm.Page.getAttribute('im_propositiontype').setValue(930760001);");
            js.ExecuteScript("Xrm.Page.getAttribute('im_propositionsubtype').setValue(930760011);");
            js.ExecuteScript("var lookupValue = new Array();lookupValue[0] = new Object();lookupValue[0].id = '{060A529E-D746-E811-814D-000D3AA27A22}';lookupValue[0].name = 'Arindam Saha';lookupValue[0].entityType = 'systemuser';Xrm.Page.getAttribute('im_salespersonid').setValue(lookupValue)");
            js.ExecuteScript("var lookupValue = new Array();lookupValue[0] = new Object();lookupValue[0].id = '{2FCCE0FA-2CBD-E611-80F3-000D3AA17340}';lookupValue[0].name = 'Customer10';lookupValue[0].entityType = 'account';Xrm.Page.getAttribute('im_headquarteraccountid').setValue(lookupValue)");

            driver.SwitchTo().ParentFrame();
            Thread.Sleep(1000);
            var save = driver.FindElement(By.XPath("//*[@id='im_project|NoRelationship|Form|Mscrm.Form.im_project.Save']/span/a/span"));
            save.Click();
        }

    }
}
