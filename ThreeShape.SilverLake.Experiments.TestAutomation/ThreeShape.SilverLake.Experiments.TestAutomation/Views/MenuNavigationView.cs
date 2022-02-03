using OpenQA.Selenium;

namespace ThreeShape.SilverLake.Experiments.TestAutomation.Views
{
    public class MenuNavigationView : BaseView
    {
        public bool IsCaseEntryPageOpened(out string actualTitle)
        {
            actualTitle = Browser.Title;
            return actualTitle.Contains("Case Flow");
        }

        public MenuNavigationView ClickCaseFlowMenu()
        {
            Browser.FindElementById("mainmenu_case_flow").Click();

            return this;
        }

        public MenuNavigationView ClickCaseEntry()
        {
            Browser.FindElementByClassName("client_tmi").Click();

            Browser.SwitchTo().Frame(Browser.FindElement(By.XPath("//iframe[@src='/ui/CaseEntry']")));

            return this;
        }

        public MenuNavigationView ClickManufacturingManager()
        {
            Browser.FindElementByXPath("//*[@id='manufacturing_manager']/span").Click();

            return this;
        }

        public MenuNavigationView SearchCaseByRefNo(string refNo)
        {
            Browser.FindElementByXPath("//*[@id='searchbox-input']").Click();
            Browser.FindElementByXPath("//*[@id='searchbox-input']").SendKeys(refNo);
            Browser.FindElementByXPath("//*[@id='searchbox-input']").SendKeys(Keys.Enter);

            return this;
        }
    }
}
