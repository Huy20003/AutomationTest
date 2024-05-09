using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject2.NhaTuyenDung.Login
{
    public class HomeLogin
    {
        private IWebDriver driver;
        private By _Link = By.CssSelector(".header__login--item.employer a[href='/nha-tuyen-dung']");
        private By _Email = By.Id("Email");
        private By _Password = By.Id("Password");
        private By _BtnLogin = By.ClassName("form__input--submit");


        public HomeLogin(IWebDriver driver)
        {
            this.driver = driver;
        }

        public HomePage Login(string email, string password)
        {
            Thread.Sleep(4000);
            driver.FindElement(_Link).Click();
            // enter email
            Thread.Sleep(4000);
            driver.FindElement(_Email).SendKeys(email);
            // enter password
            driver.FindElement(_Password).SendKeys(password);
            // click login button
            Thread.Sleep(4000);
            driver.FindElement(_BtnLogin).Click();
            // return home page
            Thread.Sleep(4000);
            return new HomePage(driver);
        }
    }
}
