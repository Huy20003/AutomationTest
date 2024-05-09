using OfficeOpenXml;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestProject2.NhaTuyenDung.TaoBaiViet;

namespace TestProject2.Admin
{
    public class DuyetPost
    {
        private IWebDriver _driver;


        [SetUp]
        public void Setup()
        {
            _driver = new ChromeDriver();
            _driver.Navigate().GoToUrl("http://localhost:62536/Admin/Login");

        }


        [TearDown]
        public void TearDown()
        {
            _driver.Quit();
        }


        public static Data[] CreateTestCase()
        {

            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<Data> testCases = new List<Data>();
            Data item = new Data();
            using (var package = new ExcelPackage(new FileInfo("C:\\Users\\ngoda\\Desktop\\TestDoAN.xlsx")))
            {
                // Chọn trang tính muốn cập nhật.
                var worksheet = package.Workbook.Worksheets["Sheet1"];

                if (worksheet == null)
                {
                    throw new Exception("Không tìm thấy trang tính");
                }
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;
                // Duyệt qua các dòng, bắt đầu từ dòng thứ 2 nếu dòng đầu tiên là tiêu đề
                for (int row = 2; row <= rowCount && row <= 10; row += 3)
                {
                    item = new Data();

                    item._title = (string)worksheet.Cells[row, 1].Value;
                    item._image = (string)worksheet.Cells[row + 1, 1].Value;
                    item._content = (string)worksheet.Cells[row + 2, 1].Value;
                    item._expectedResult = (string)worksheet.Cells[row, 2].Value;
                    item.row = row;
                    item.col = 3;



                    testCases.Add(item);
                }
            }

            return testCases.ToArray();
        }

        public void SaveExcel(string type, Data data)
        {
            using (var package = new ExcelPackage(new FileInfo("C:\\Users\\ngoda\\Desktop\\TestDoAN.xlsx")))
            {
                var worksheet = package.Workbook.Worksheets["Sheet1"];
                worksheet.Cells[data.row, data.col].Value = type;
                package.Save();
            }
        }
        [TestCaseSource(nameof(CreateTestCase))]
        public void Create(Data data)
        {
            _driver.FindElement(By.Id("UserName")).SendKeys("Admin@gmail.com");

            _driver.FindElement(By.Id("PassWord")).SendKeys("123456");


            _driver.FindElement(By.XPath("//button[text()='Đăng nhập']"));
            System.Threading.Thread.Sleep(2000);
            _driver.FindElement(By.XPath("//span[text()='Quản lý bài viết']")).Click();


            _driver.FindElement(By.CssSelector("a[href='/Admin/BaiViet/ChoDuyet'].waves-effect.active"));

            _driver.FindElement(By.CssSelector("td[style='display: inline-grid;'] a.btn.btn-success"));

            IAlert alert = _driver.SwitchTo().Alert();

            alert.Accept();


            IWebElement badgeElement = _driver.FindElement(By.CssSelector("span.badge"));
            string badgeText = badgeElement.Text;

            IWebElement alertElement = (IWebElement)_driver.FindElements(By.XPath("(//div[@id='msgAlert'])[1]"));
        }
    }
}
