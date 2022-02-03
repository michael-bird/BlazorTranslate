using ThreeShape.SilverLake.Experiments.SIL165.Drivers;
using ThreeShape.SilverLake.Experiments.SIL165.Views;
using NUnit.Framework;
using TechTalk.SpecFlow;

namespace ThreeShape.SilverLake.Experiments.SIL165.StepDefinitions
{
    [Binding]
    public sealed class StepDefinitions
    {
        private string _urlEnvVarName = "";
        private string _loginEnvVarName = "";
        private string _passEnvVarName = "";

        [Given(@"I use (.*) website")]
        public void GivenIUseWebsite(string website)
        {
            if (website.Equals("local"))
            {
                _urlEnvVarName = "url2";
                _loginEnvVarName = "login2";
                _passEnvVarName = "pass2";
            }
            else if (website.Equals("Igor"))
            {
                _urlEnvVarName = "url1";
                _loginEnvVarName = "login1";
                _passEnvVarName = "pass1";
            }
        }

        [Given(@"I have started LabStar website in Chrome")]
        public void GivenIHaveStartedChrome()
        {
            string? url = Environment.GetEnvironmentVariable(_urlEnvVarName);
            if (url == null) throw new Exception("Environmental variable for url should be added");

            ChromeBrowser.Browser.Navigate().GoToUrl(url);
        }

        [Given(@"proceed to the main page")]
        public void GivenProceedToTheMainPage()
        {
            string? login = Environment.GetEnvironmentVariable(_loginEnvVarName);
            if (login == null) throw new Exception("Environmental variable for login should be added");
            string? pass = Environment.GetEnvironmentVariable(_passEnvVarName);
            if (pass == null) throw new Exception("Environmental variable for password should be added");

            ChromeBrowser.Browser
                                .GetView<LoginPageView>()
                                .EnterUserName(login)
                                .EnterPassword(pass)
                                .ClickLoginButton();

            var res = ChromeBrowser.Browser
                .GetView<MenuNavigationView>()
                .IsCaseEntryPageOpened(out string title);

            Assert.IsTrue(res,
                $"Login failed or incorrect main page is opened. Actual page title is: {title}");
        }

        [When(@"I open Case Entry")]
        public void WhenIOpenCaseEntry()
        {
            ChromeBrowser.Browser
                .GetView<MenuNavigationView>()
                .ClickCaseFlowMenu()
                .ClickCaseEntry();

            Assert.IsTrue(ChromeBrowser.Browser
                .GetView<CaseEntryFormView>().IsCaseEntryFormOpened(), "Create Case form is not shown");
        }

        [When(@"fill the form")]
        public void WhenFillTheForm(Table table)
        {
            string doctor = "";
            string firstName = "";
            string lastName = "";
            string type = "";
            int tooth = 0;
            string item = "";
            string schedule = "";

            foreach (var row in table.Rows)
            {
                switch (row["Field"])
                {
                    case "Doctor/Lab":
                    {
                        doctor = row["Value"];
                        break;
                    }
                    case "Patient first name":
                    {
                        firstName = row["Value"];
                        break;
                    }
                    case "Patient last name":
                    {
                        lastName = row["Value"];
                        break;
                    }
                    case "Item type":
                    {
                        type = row["Value"];
                        break;
                    }
                    case "Tooth":
                    {
                        tooth = Int32.Parse(row["Value"]);
                        break;
                    }
                    case "Item":
                    {
                        item = row["Value"];
                        break;
                    }
                    case "Schedule":
                    {
                        schedule = row["Value"];
                        break;
                    }
                    default:
                    {
                        Console.WriteLine($"Incorrect field {row["Field"]} is added to scenario, so it is ignored.");
                        break;
                    }
                }
            }
            
            ChromeBrowser.Browser
                .GetView<CaseEntryFormView>()
                .ChooseDoctor(doctor)
                .FillFirstName(firstName)
                .FillLastName(lastName)
                .ChooseTooth(tooth)
                .ChooseType(type)
                .ChooseItem(item)
                .ChooseSchedule(schedule);
        }

        [When(@"click Create button")]
        public void WhenClickCreateButton()
        {
            ChromeBrowser.Browser.GetView<CaseEntryFormView>().ClickCreateButton();
        }

        [Then(@"case should be created and visible in the Manufacturing Manager")]
        public void ThenCaseShouldBeCreatedAndVisibleInTheManufacturingManager()
        {
            ChromeBrowser.Browser
                .GetView<CaseEntryFormView>()
                .IsNextStepsFormShown()
                .IsWorkticketShown()
                .ViewCaseRecord(out string refNo)
                .CloseCaseRecordForm();
            ChromeBrowser.Browser
                .GetView<MenuNavigationView>()
                .ClickManufacturingManager()
                .SearchCaseByRefNo(refNo);
            ChromeBrowser.Browser
                .GetView<ManufacturingManagerView>()
                .DoesCaseExists();
        }

        [AfterScenario()]
        public void CloseDriver()
        {
            ChromeBrowser.Browser.Close();
        }
    }
}