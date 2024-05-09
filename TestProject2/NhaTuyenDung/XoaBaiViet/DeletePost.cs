using OfficeOpenXml;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestProject2.NhaTuyenDung.TaoBaiViet;
using System.Collections.ObjectModel;
using TestProject2.NhaTuyenDung.Login;

namespace TestProject2.NhaTuyenDung.XoaBaiViet
{
    public class DeletePost
    {
        private IWebDriver _driver;


        [SetUp]
        public void Setup()
        {
            _driver = new ChromeDriver();
            _driver.Navigate().GoToUrl("http://localhost:62536/Home/Index");
            _driver.Manage().Window.Maximize();

        }


        [TearDown]
        public void TearDown()
        {
            _driver.Quit();
        }


        public static Data2[] CreateTestCase()
        {

            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<Data2> testCases1 = new List<Data2>();
            Data2 item = new Data2();
            using (var package = new ExcelPackage(new FileInfo("C:\\Users\\ngoda\\Desktop\\TestDoAN.xlsx")))
            {
                // Chọn trang tính muốn cập nhật.
                var worksheet = package.Workbook.Worksheets["Sheet3"];

                if (worksheet == null)
                {
                    throw new Exception("Không tìm thấy trang tính");
                }
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;
                // Duyệt qua các dòng, bắt đầu từ dòng thứ 2 nếu dòng đầu tiên là tiêu đề
                for (int row = 2; row <= rowCount && row <= 5; row += 1)
                {
                    item = new Data2();

                    item._MoTa = (string)worksheet.Cells[row, 1].Value;
                    item._ExpectedResult = (string)worksheet.Cells[row, 2].Value;

                    item.row = row;
                    item.col = 3;



                    testCases1.Add(item);
                }
            }

            return testCases1.ToArray();
        }

        public void SaveExcel(string type, Data2 data)
        {
            using (var package = new ExcelPackage(new FileInfo("C:\\Users\\ngoda\\Desktop\\TestDoAN.xlsx")))
            {
                var worksheet = package.Workbook.Worksheets["Sheet3"];
                worksheet.Cells[data.row, data.col].Value = type;
                package.Save();
            }
        }

        [TestCaseSource(nameof(CreateTestCase))]
        public void Create(Data2 data)
        {
            HomeLogin homeLogin = new HomeLogin(_driver);
            HomePage homepage = homeLogin.Login("Phat@gmail.com", "123456");

            Thread.Sleep(2000);
            string save = "Fail";
            _driver.FindElement(By.CssSelector("li:nth-child(5) > .waves-effect")).Click();
            Thread.Sleep(2000);
            _driver.FindElement(By.CssSelector("a[href='/nha-tuyen-dung/BaiViet/ChoDuyet']")).Click();

            Thread.Sleep(2000);




            if (data._MoTa.Equals("Xóa bài viết mã 19 thành công"))
            {
                // Tìm phần tử <a> chứa số 2 trong danh sách trang
                IWebElement page2Element = _driver.FindElement(By.CssSelector("#load-pagination [data-page='2']"));

                // Click vào phần tử số 2
                page2Element.Click();
                Thread.Sleep(2000);

                _driver.FindElement(By.CssSelector("td.d-flex a.btn-danger")).Click();

                IAlert alert = _driver.SwitchTo().Alert();
                string alertText = alert.Text;
                Thread.Sleep(2000);

                if (data._ExpectedResult.Equals(alertText))
                {
                    alert.Accept();
                    save = "Pass";
                    SaveExcel(save, data);
                }
                else
                {
                    save = "Fail";
                    SaveExcel(save, data);
                }

            }
            else if (data._MoTa.Equals("Bấm hủy không xóa bài viết"))
            {
                var spanElement = _driver.FindElement(By.XPath("//li[@class='mm-active']/a/span[@class='badge badge-pink badge-pill float-right']"));
                Thread.Sleep(2000);

                string text = spanElement.Text;
                Thread.Sleep(2000);

                // Tìm phần tử <a> chứa số 2 trong danh sách trang
                IWebElement page2Element = _driver.FindElement(By.CssSelector("#load-pagination [data-page='2']"));
                Thread.Sleep(2000);

                // Click vào phần tử số 2
                page2Element.Click();
                Thread.Sleep(2000);

                _driver.FindElement(By.CssSelector("td.d-flex a.btn-danger")).Click();

                Thread.Sleep(2000);

                IAlert alert = _driver.SwitchTo().Alert();
                alert.Dismiss();
                Thread.Sleep(2000);

                var spanElement2 = _driver.FindElement(By.XPath("//li[@class='mm-active']/a/span[@class='badge badge-pink badge-pill float-right']"));
                string text2 = spanElement2.Text;
                if (text.Equals(text2))
                {
                    save = "Pass";
                    SaveExcel(save, data);
                }
                else
                {
                    save = "Fail";
                    SaveExcel(save, data);
                }
            }
            else if (data._MoTa.Equals("Xóa bài viết,kiểm tra số lượng bài viết\r\ncó bị giảm đi không"))
            {
                var spanElement3 = _driver.FindElement(By.XPath("//li[@class='mm-active']/a/span[@class='badge badge-pink badge-pill float-right']"));
                string text3 = spanElement3.Text;
                Thread.Sleep(2000);

                // Tìm phần tử <a> chứa số 2 trong danh sách trang
                IWebElement page2Element = _driver.FindElement(By.CssSelector("#load-pagination [data-page='2']"));
                Thread.Sleep(2000);

                // Click vào phần tử số 2
                page2Element.Click();
                Thread.Sleep(2000);

                _driver.FindElement(By.CssSelector("td.d-flex a.btn-danger")).Click();
                Thread.Sleep(2000);

                IAlert alert = _driver.SwitchTo().Alert();
                alert.Accept();
                Thread.Sleep(2000);

                var spanElement4 = _driver.FindElement(By.XPath("//li[@class='mm-active']/a/span[@class='badge badge-pink badge-pill float-right']"));
                string text4 = spanElement4.Text;
                if (text3.CompareTo(text4) > 0)
                {
                    save = "Pass";
                    SaveExcel(save, data);
                }
                else
                {
                    save = "Fail";
                    SaveExcel(save, data);
                }

            }
            else if (data._MoTa.Equals("Xóa bài viết ,tìm thử bài viết đã xóa"))
            {
                Thread.Sleep(2000);

                IWebElement searchBox = _driver.FindElement(By.Id("txtsearch"));
                searchBox.SendKeys("HUYYYYYYYYádwqd");
                Thread.Sleep(2000);

                ReadOnlyCollection<IWebElement> tenBaiVietElements = _driver.FindElements(By.XPath("//table[@id='datatablesSimple']//tbody//td[3]"));
                Thread.Sleep(2000);

                foreach (IWebElement tenBaiVietElement in tenBaiVietElements)
                {
                    string tenBaiViet = tenBaiVietElement.Text;
                    if (tenBaiViet.Equals("HUYYYYYYYYádwqd"))
                    {
                        save = "Pass";
                        SaveExcel(save, data);
                        break;
                    }
                    else
                    {
                        save = "Fail";
                        SaveExcel(save, data);
                        break;
                    }
                }

            }




            // Trong phần tử <tr>, tìm phần tử <a> có thuộc tính data-user="8" (nút Xóa)




            // Đóng thông báo xác nhận


        }
    }
}
