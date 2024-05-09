using OfficeOpenXml;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestProject2.NhaTuyenDung.TaoBaiViet;
using TestProject2.NhaTuyenDung.Login;

namespace TestProject2.NhaTuyenDung.SuaBaiViet
{
    public class EditPost
    {
        private IWebDriver _driver;


        [SetUp]
        public void Setup()
        {
            _driver = new ChromeDriver();
            _driver.Navigate().GoToUrl("http://localhost:62536/Home/Index");

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
                var worksheet = package.Workbook.Worksheets["Sheet2"];

                if (worksheet == null)
                {
                    throw new Exception("Không tìm thấy trang tính");
                }
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;
                // Duyệt qua các dòng, bắt đầu từ dòng thứ 2 nếu dòng đầu tiên là tiêu đề
                for (int row = 2; row <= rowCount && row <= 16; row += 3)
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
                var worksheet = package.Workbook.Worksheets["Sheet2"];
                worksheet.Cells[data.row, data.col].Value = type;
                package.Save();
            }
        }

        [TestCaseSource(nameof(CreateTestCase))]
        public void Create(Data data)
        {
            HomeLogin homeLogin = new HomeLogin(_driver);
            HomePage homepage = homeLogin.Login("Phat@gmail.com", "123456");

            _driver.FindElement(By.CssSelector("li:nth-child(5) > .waves-effect")).Click();
            Thread.Sleep(2000);
            _driver.FindElement(By.CssSelector("a[href='/nha-tuyen-dung/BaiViet/ChoDuyet']")).Click();

            _driver.FindElement(By.XPath("//td[@class='d-flex']//a[contains(@class, 'btn-warning')]")).Click();
            // Trong phần tử <td>, tìm phần tử <a> có văn bản là "Sửa"

            // Không có thông báo xuất hiện, tiếp tục thực hiện các hoạt động khác
            if (!string.IsNullOrEmpty(data._title))
            {
                IWebElement tilte = _driver.FindElement(By.Id("TenBaiViet"));
                tilte.Clear();
                tilte.SendKeys(data._title);
            }
            else
            {
                IWebElement tilte = _driver.FindElement(By.Id("TenBaiViet"));
                tilte.Clear();
            }
            _driver.FindElement(By.Name("Image")).SendKeys(data._image);
            Thread.Sleep(2000);

            if (!string.IsNullOrEmpty(data._content))
            {
                IWebElement iframe = _driver.FindElement(By.CssSelector("iframe.cke_wysiwyg_frame"));
                _driver.SwitchTo().Frame(iframe);
                IWebElement editorBody = _driver.FindElement(By.CssSelector("body.cke_editable"));
                editorBody.SendKeys(Keys.Control + "a"); // Chọn toàn bộ nội dung
                editorBody.SendKeys(Keys.Delete);
                //editorBody.Click();
                editorBody.SendKeys(data._content);
                _driver.SwitchTo().DefaultContent();
            }
            else
            {
                IWebElement iframe = _driver.FindElement(By.CssSelector("iframe.cke_wysiwyg_frame"));
                _driver.SwitchTo().Frame(iframe);
                IWebElement editorBody = _driver.FindElement(By.CssSelector("body.cke_editable"));
                editorBody.SendKeys(Keys.Control + "a");
                editorBody.SendKeys(Keys.Delete);
                _driver.SwitchTo().DefaultContent();
            }
            Thread.Sleep(2000);


            _driver.FindElement(By.CssSelector("input[type='submit'][value='Cập nhật']")).Click();
            Thread.Sleep(2000);

            string save = "Fail";

            if (data._expectedResult.Equals("Cập nhật thành công"))
            {
                IWebElement successAlert = _driver.FindElement(By.XPath("//div[@class='alert alert-success' and contains(text(),'Cập nhật thành công')]"));

                // Lấy nội dung của thông báo
                string successMessage = successAlert.Text;
                if (data._expectedResult.Equals(successMessage))
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
            else if (data._expectedResult.Equals("Cập nhật không thành công"))
            {

                IWebElement successAlert = _driver.FindElement(By.XPath("//div[@class='alert alert-success' and contains(text(),'Cập nhật thành công')]"));

                // Lấy nội dung của thông báo
                string successMessage = successAlert.Text;
                if (data._expectedResult.Equals(successMessage))
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
            else if (data._expectedResult.Equals("Bạn chưa nhập tên bài viết"))
            {

                var validationMessageElement = _driver.FindElement(By.Id("TenBaiViet-error"));
                string validationMessage = validationMessageElement.Text;
                if (data._expectedResult.Equals(validationMessage))
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
            else if (data._expectedResult.Equals("Tên bài viết đã được thay đổi"))
            {
                _driver.FindElement(By.CssSelector("a[href='/nha-tuyen-dung/BaiViet/ChoDuyet']")).Click();

                _driver.FindElement(By.XPath("//a[text()='2']"));

                var textName = _driver.FindElement(By.XPath("//td[contains(text(),'Lập trình thiết bị di động')]"));
                string text = textName.Text;
                if (data._title.Equals(text))
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
            else if (data._expectedResult.Equals("Cập nhật thất bại"))
            {

                var validationMessageElement3 = _driver.FindElement(By.CssSelector("span.field-validation-error.text-danger"));
                string validationMessage3 = validationMessageElement3.Text;
                if (validationMessage3.Equals("Bạn chưa nhập nội dung"))
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

        }
    }
}
