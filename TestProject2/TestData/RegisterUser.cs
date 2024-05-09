using OfficeOpenXml;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using TestProject2.NhaTuyenDung.TaoBaiViet;

using System.Data.Entity;
using DoAn2023.Models.EF;
using System.Data.SqlClient;


namespace TestProject2.TestData
{
    public class RegisterUser
    {
        private IWebDriver _driver;


        private const string ConnectionString = "data source=DESKTOP-QV8266K\\SQLEXPRESS;initial catalog=Career;Trusted_Connection=True;";

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

        public static DataRegister[] CreateTestCase()
        {

            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<DataRegister> testCases5 = new List<DataRegister>();
            DataRegister item = new DataRegister();
            using (var package = new ExcelPackage(new FileInfo("C:\\Users\\ngoda\\Desktop\\TestDoAN.xlsx")))
            {
                // Chọn trang tính muốn cập nhật.
                var worksheet = package.Workbook.Worksheets["Sheet4"];

                if (worksheet == null)
                {
                    throw new Exception("Không tìm thấy trang tính");
                }
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;
                // Duyệt qua các dòng, bắt đầu từ dòng thứ 2 nếu dòng đầu tiên là tiêu đề
                for (int row = 2; row <= rowCount && row <= 5; row += 4)
                {
                    item = new DataRegister();

                    item._FullName = (string)worksheet.Cells[row, 1].Value;
                    item._email = (string)worksheet.Cells[row + 1, 1].Value;
                    item._passWord = (string)worksheet.Cells[row + 2, 1].Value.ToString();
                    item._repassword = (string)worksheet.Cells[row+3, 1].Value.ToString();
                    item.row = row;
                    item.col = 3;



                    testCases5.Add(item);
                }
            }

            return testCases5.ToArray();
        }

        public void SaveExcel(string type, DataRegister data)
        {
            using (var package = new ExcelPackage(new FileInfo("C:\\Users\\ngoda\\Desktop\\TestDoAN.xlsx")))
            {
                var worksheet = package.Workbook.Worksheets["Sheet4"];
                worksheet.Cells[data.row, data.col].Value = type;
                package.Save();
            }
        }
        [TestCaseSource(nameof(CreateTestCase))]
        public void Create( DataRegister data)
        {
            string save = "Fail";
            _driver.FindElement(By.Id("btn-register")).Click();

            IWebElement namelElement = _driver.FindElement(By.Id("register_name"));
            namelElement.SendKeys(data._FullName);

          
            Thread.Sleep(2000);

            IWebElement emailElement = _driver.FindElement(By.Id("register_email"));
               emailElement.SendKeys(data._email);
            var emailValue = emailElement.GetAttribute("value");
            Thread.Sleep(2000);
            _driver.FindElement(By.Id("register_password")).SendKeys(data._passWord);
            Thread.Sleep(2000);
            _driver.FindElement(By.Id("password_confirm")).SendKeys(data._repassword);

            Thread.Sleep(3000);
            IWebElement registerButton = _driver.FindElement(By.Id("form_btn-register"));

            // Thực hiện click vào button
            registerButton.Click();
            Thread.Sleep(3000);
            string emailToCheck = emailValue;
            string connectionString = "data source=DESKTOP-QV8266K\\SQLEXPRESS;initial catalog=Career;Trusted_Connection=True;";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = "SELECT COUNT(*) FROM tbl_TaiKhoan WHERE sEmail = @Email";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Email", emailToCheck);

                    int count = (int)command.ExecuteScalar();
                    if (count > 0)
                    {
                        save = "Pass";
                        SaveExcel(save,data);
                    }
                    else
                    {
                        save = "Fail";
                        SaveExcel(save, data);
                    }
                }
            }


        }
       
        //public void CanReadDataFromTestDatabase()
        //{
        //    int rowCount = 0;

        //    // Kết nối đến test database
        //    using (var connection = new SqlConnection(ConnectionString))
        //    {
        //        connection.Open();

        //        // Thực hiện truy vấn đến test database
        //        using (var command = new SqlCommand("SELECT COUNT(*) FROM TestTable", connection))
        //        {
        //            rowCount = (int)command.ExecuteScalar();
        //        }
        //    }

        //    // Kiểm tra số lượng hàng trả về
        //    Assert.Equal(3, rowCount);
        //}

    }

}
