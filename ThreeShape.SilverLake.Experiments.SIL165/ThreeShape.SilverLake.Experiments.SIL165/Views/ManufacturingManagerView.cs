using NUnit.Framework;
using OpenQA.Selenium;

namespace ThreeShape.SilverLake.Experiments.SIL165.Views
{
    public class ManufacturingManagerView : BaseView
    {
        public ManufacturingManagerView DoesCaseExists()
        {
            //Wait untill table is shown
            Drivers.WebDriverExtensions.FindElement(Browser, By.XPath("//*[@id='tabcontent-manu']/div[2]/table/tbody"));

            int searchRes = Browser.FindElementsByXPath("//*[@id='tabcontent-manu']/div[2]/table/tbody/tr").Count;
            Assert.IsTrue(searchRes == 1, $"Incorrect amount of cases found: {searchRes}");

            return this;
        }
    }
}
