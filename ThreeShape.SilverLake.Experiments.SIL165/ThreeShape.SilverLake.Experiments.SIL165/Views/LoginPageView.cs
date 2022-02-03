namespace ThreeShape.SilverLake.Experiments.SIL165.Views
{
    public class LoginPageView : BaseView
    {
        public LoginPageView EnterUserName(string login)
        {
            Browser.FindElementById("username").SendKeys(login);
            return this;
        }

        public LoginPageView EnterPassword(string pass)
        {
            Browser.FindElementById("password").SendKeys(pass);
            return this;
        }

        public LoginPageView ClickLoginButton()
        {
            Browser.FindElementById("Submit1").Click();
            return this;
        }
    }
}
