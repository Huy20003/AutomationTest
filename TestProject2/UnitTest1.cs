using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.ComponentModel.DataAnnotations;
using System.Web;






namespace TestProject2
  
{
    public class Tests
    {
        public IWebDriver driver;
        string _btnLogin = "btn-login";
        string _userLogin = "login_email";
        string _userPass = "login_password";
        string _submitLogin = "form_btn-login";
        By _message = By.XPath("//div[@id='msgAlert']");
        By _btnRegister = By.Id("btn-register");
        By _nameRegister = By.Id("register_name");
        By _emailRegister = By.Id("register_email");
        By _passRegister = By.Id("register_password");
        By _rePass = By.Id("password_confirm");
        By _submitRegis = By.Id("form_btn-register");
        By _nameUser = By.ClassName("user__info--name");
          public   string _nameValue="";
        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("http://localhost:62536/Home/Index");

        }

        [TearDown]
        public void TearDown()
        {
            driver.Quit();
        }

        //[Test]
        //[TestCase("Tran Tuong Huan", "huan@gmail.com", "12345", "12345")]   //đổi dữ liệu chỗ này
        //public void TestRegister(string name, string email, string pass, string repass)
        //{
        //    driver.FindElement(_btnRegister).Click();
        //    Thread.Sleep(2000);
        //    driver.FindElement(_nameRegister).SendKeys(name);
        //    _nameValue = name;

        //    Thread.Sleep(2000);
        //    driver.FindElement(_emailRegister).SendKeys(email);
        //    Thread.Sleep(2000);
        //    driver.FindElement(_passRegister).SendKeys(pass);
        //    Thread.Sleep(2000);
        //    driver.FindElement(_rePass).SendKeys(repass);
        //    Thread.Sleep(2000);
        //    driver.FindElement(_submitRegis).Click();
        //    Thread.Sleep(2000);

        //    _nameValue = GetNameRegisterValue();

        //}

        public string GetNameRegisterValue()
        {
            return driver.FindElement(_nameRegister).GetAttribute("value");
        }

        [Test]
        [TestCase("huan@gmail.com","12345")]
        public void TestLogin(string email, string pass)
        {
            driver.FindElement(By.Id(_btnLogin)).Click();
            Thread.Sleep(2000);

            driver.FindElement(By.Id(_userLogin)).SendKeys(email);
            Thread.Sleep(2000);

            driver.FindElement(By.Id(_userPass)).SendKeys(pass);
            Thread.Sleep(2000);

            driver.FindElement(By.Id(_submitLogin)).Submit();
            Thread.Sleep(2000);


            var actualResult = driver.FindElement(_message);
            string txt = actualResult.Text;
            if(string.IsNullOrWhiteSpace(txt)) { txt = actualResult.GetAttribute("value"); }
            if (string.IsNullOrWhiteSpace(txt)) { txt = actualResult.GetAttribute("innerText"); }
            if (string.IsNullOrWhiteSpace(txt)) { txt = actualResult.GetAttribute("textContent"); }
            Assert.IsTrue(actualResult.Text.Equals("Đăng nhập thành công !"));
            //// này đg bị lỗi 


            //var actualName = driver.FindElement(_nameUser);
            //Assert.IsTrue(actualName.Text.Equals(_nameValue));

        }
       

    }
}