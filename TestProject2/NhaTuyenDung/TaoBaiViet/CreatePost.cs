using Microsoft.VisualStudio.TestPlatform.CommunicationUtilities;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using TestProject2.NhaTuyenDung.Login;
using System.Data.SqlClient;
using System.Drawing;



namespace TestProject2.NhaTuyenDung.TaoBaiViet
{   
    public class CreatePost
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
            HomeLogin homeLogin = new HomeLogin(_driver);
            HomePage homepage = homeLogin.Login("Phat@gmail.com", "123456");

            try
            {
                IAlert alert = _driver.SwitchTo().Alert();
                alert.Accept();
            }
            catch (NoAlertPresentException)
            {
                _driver.FindElement(By.CssSelector("li:nth-child(5) > .waves-effect")).Click();
                Thread.Sleep(4000);
                _driver.FindElement(By.CssSelector("a[href='/nha-tuyen-dung/BaiViet/Create']")).Click();
                // Không có thông báo xuất hiện, tiếp tục thực hiện các hoạt động khác
                if (!string.IsNullOrEmpty(data._title))
                {
                    IWebElement ten = _driver.FindElement(By.Id("TenBaiViet"));
                        ten.SendKeys(data._title);
                  
                var emailValue = ten.GetAttribute("value");
                string tenCheck = emailValue;
                }
                _driver.FindElement(By.Name("Image")).SendKeys(data._image);
                Thread.Sleep(4000);
                if (!string.IsNullOrEmpty(data._content))
                {
                    IWebElement iframe = _driver.FindElement(By.CssSelector("iframe.cke_wysiwyg_frame"));
                    _driver.SwitchTo().Frame(iframe);
                    IWebElement editorBody = _driver.FindElement(By.CssSelector("body.cke_editable"));
                    editorBody.Click();
                    editorBody.SendKeys(data._content);
                    _driver.SwitchTo().DefaultContent();
                }
                Thread.Sleep(4000);

                _driver.FindElement(By.CssSelector("input[type='submit'][value='Tạo mới']")).Click();

                Thread.Sleep(3000);

                //IWebElement errorMessage = _driver.FindElement(By.CssSelector("span#TenBaiViet-error"));
                //string text=errorMessage.Text;


                IWebElement ten1 = _driver.FindElement(By.Id("TenBaiViet"));
                ten1.SendKeys(data._title);

                var tenValue1 = ten1.GetAttribute("value");
                string tenCheck1 = tenValue1;


                //var actualResult = driver.FindElement(_message);
                //string txt = actualResult.Text;
                //if (string.IsNullOrWhiteSpace(validationMessage)) { validationMessage = validationMessageElement.GetAttribute("value"); }
                //if (string.IsNullOrWhiteSpace(validationMessage)) { validationMessage = validationMessageElement.GetAttribute("innerText"); }
                //if (string.IsNullOrWhiteSpace(validationMessage)) { validationMessage = validationMessageElement.GetAttribute("textContent"); }



                //string _successMessage = _driver.FindElement(By.Id("msgAlert")).Text;

                bool status = false;
                string save = "Fail";
                if (data._expectedResult.Equals("Bạn chưa nhập tên bài viết"))
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
                else if (data._expectedResult.Equals("Tên bài viết đã tồn tại"))
                {
                    var validationMessageElement2 = _driver.FindElement(By.CssSelector(".validation-summary-errors ul li"));
                    string validationMessage2 = validationMessageElement2.Text;
                    if (data._expectedResult.Equals(validationMessage2))
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
                else if (data._expectedResult.Equals("Bạn chưa nhập nội dung"))
                {
                    var validationMessageElement3 = _driver.FindElement(By.CssSelector("span.field-validation-error.text-danger"));
                    string validationMessage3 = validationMessageElement3.Text;
                    if (data._expectedResult.Equals(validationMessage3))
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
                else if (data._expectedResult.Equals("Tạo bài viết thành công"))
                {
                    var successMessgae = _driver.FindElement(By.Id("msgAlert"));
                    string SuccesMessage = successMessgae.Text;
                    if (data._expectedResult.Equals(SuccesMessage))
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
                else if (data._expectedResult.Equals("Lưu trong database thành công"))
                    {

                    var successMessgae5 = _driver.FindElement(By.Id("msgAlert"));
                    string SuccesMessage4 = successMessgae5.Text;

                    if(SuccesMessage4.Equals("Tạo bài viết thành công"))
                    {
                        string connectionString = "data source=DESKTOP-QV8266K\\SQLEXPRESS;initial catalog=Career;Trusted_Connection=True;";

                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            connection.Open();

                            string query = "SELECT COUNT(*) FROM tbl_TaiKhoan WHERE sTenBaiViet = @sTenBaiViet";
                            using (SqlCommand command = new SqlCommand(query, connection))
                            {
                                command.Parameters.AddWithValue("@sTenBaiViet", tenCheck1);

                                int count = (int)command.ExecuteScalar();
                                if (count > 0)
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
                   


                //else if(data._expectedResult.Equals(_errorMessage1))








            }

            }
        }
    }

