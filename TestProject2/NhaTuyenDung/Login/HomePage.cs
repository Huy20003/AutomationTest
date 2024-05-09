using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject2.NhaTuyenDung.Login
{
    public class HomePage
    {
        private IWebDriver driver;
        private By _successLogin = By.XPath("//div[@id='msgAlert']");

        public HomePage(IWebDriver driver)
        {
            this.driver = driver;
        }

        public bool IsDisplayed()
        {
            if (Utils.WaitForElementDisplay(driver, _successLogin, 10))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
