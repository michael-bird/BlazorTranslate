using ThreeShape.SilverLake.Experiments.SIL165.Drivers;
using NUnit.Framework;
using OpenQA.Selenium;

namespace ThreeShape.SilverLake.Experiments.SIL165.Views
{
    public class CaseEntryFormView : BaseView
    {
        public bool IsCaseEntryFormOpened()
        {
            return Browser.FindElementsByXPath("//div[@class='header-title'][contains(text(),'Create Case')]").Count == 1;
        }

        public CaseEntryFormView ChooseDoctor(string doctor)
        {
            WebDriverExtensions.FindElement(Browser, By.XPath("//*[@id='case-upsert-doctor-or-lab']/div/span/span/span[2]/span")).Click();
            WebDriverExtensions.FindElement(Browser, By.XPath("//*[contains(text(),'" + doctor + "')]")).Click();

            return this;
        }

        public CaseEntryFormView FillFirstName(string firstName)
        {
            Browser.FindElementByXPath("//div[@class='patient-name-checking-first-name']/div/input").SendKeys(firstName);

            return this;
        }

        public CaseEntryFormView FillLastName(string lastName)
        {
            Browser.FindElementByXPath("//div[@class='patient-name-checking-last-name']/div/input").SendKeys(lastName);

            return this;
        }

        public CaseEntryFormView ChooseTooth(int toothN)
        {
            //For local
            Browser.FindElementByXPath("//*[@id='itemSelection']/div[2]/table/tbody/tr/td[2]/button/span").Click();
            //Fro Igor
            //Browser.FindElementByXPath("//*[@id='itemSelection']/div[2]/table/tbody/tr[1]/td[3]/button/span").Click();
            
            Browser.FindElementById("tooth" + toothN + "_wrapper").Click();

            return this;
        }

        public CaseEntryFormView ChooseType(string type)
        {
            //For local
            WebDriverExtensions.FindElement(Browser, By.XPath("//*[@id='itemSelection']/div[2]/table/tbody/tr[1]/td[3]/button/div")).Click();
            //For Igor
            //WebDriverExtensions.FindElement(Browser, By.XPath("//*[@id='itemSelection']/div[2]/table/tbody/tr[1]/td[2]/button/div")).Click();
            
            WebDriverExtensions.FindElement(Browser, By.XPath("//*[contains(text(),'" + type + "')]")).Click();
            
            return this;
        }

        public CaseEntryFormView ChooseItem(string item)
        {
            //For local
            WebDriverExtensions.FindElement(Browser, By.XPath("//*[@id='itemSelection']/div[2]/table/tbody/tr[1]/td[4]/button")).Click();
            //For Igor
            //WebDriverExtensions.FindElement(Browser, By.XPath("//*[@id='itemSelection']/div[2]/table/tbody/tr[1]/td[4]/button")).Click();
            
            WebDriverExtensions.FindElement(Browser, By.XPath("//*[contains(text(),'" + item + "')]")).Click();
            
            return this;
        }

        public CaseEntryFormView ChooseSchedule(string schedule)
        {
            //For local
            WebDriverExtensions.FindElement(Browser, By.XPath("//*[@id='root']/div[2]/div[6]/div[2]/div[2]/div[1]/div/span/span/span[2]")).Click();
            //For Igor
            //WebDriverExtensions.FindElement(Browser, By.XPath("//*[@id='root']/div[2]/div[6]/div[2]/div[3]/div[1]/div/span/span/span[2]")).Click();
            
            WebDriverExtensions.FindElement(Browser, By.XPath("//*[contains(text(),'" + schedule + "')]")).Click();
            
            return this;
        }
        
        public CaseEntryFormView ClickCreateButton()
        {
            Browser.FindElementById("caseSaveButton").Click();

            return this;
        }

        public CaseEntryFormView IsNextStepsFormShown()
        {
            IWebElement form = Browser.FindElementByXPath("//h1[contains(text(), 'Next steps')]");
            Assert.IsTrue(form.Displayed && form.Enabled, $"Next steps form is either not displayed ({form.Displayed}) or enabled ({form.Enabled})");

            return this;
        }

        public CaseEntryFormView IsWorkticketShown()
        {
            var mainWindowHandle = Browser.CurrentWindowHandle;
            Browser.SwitchTo().Window(Browser.WindowHandles.Last());
            
            string url = Browser.Url;
            Assert.IsTrue(url.Contains("pdf") && url.Contains("workticket"), $"Incorrect url: {url}");

            int pdfAmount = Browser.FindElementsByXPath("//embed[@type='application/pdf']").Count;
            Assert.IsTrue(pdfAmount == 1, $"There is no or more than 1 embedded pdf files: {pdfAmount}");
            
            //Browser.GetScreenshot().SaveAsFile("C:\\Automation\\Screenshots\\Screen1.png");

            Browser.Close();
            Browser.SwitchTo().Window(mainWindowHandle);

            return this;
        }

        public CaseEntryFormView ViewCaseRecord(out string refNo)
        {
            Browser.SwitchTo().Frame(Browser.FindElement(By.XPath("//iframe[@src='/ui/CaseEntry']")));

            Browser.FindElementByXPath("//*[contains(text(),'View Case Record')]").Click();
            refNo = Browser.FindElementByXPath("//*[@id='root']/div[2]/div/header/div/div[1]/span[2]/strong").Text;

            return this;
        }

        public CaseEntryFormView CloseCaseRecordForm()
        {
            Browser.SwitchTo().ParentFrame();
            Browser.FindElementByXPath("//*[@id='ext-gen15']/div[13]/div[1]/div/a").Click();

            return this;
        }
    }
}
