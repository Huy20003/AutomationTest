using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject2.NhaTuyenDung.Login
{
    public class Utils
    {
        public static bool WaitForElementDisplay(IWebDriver driver, By by, int waitInSeconds)
        {
            for (int i = 0; i < waitInSeconds / 2 + 1; i++)
            {
                try
                {
                    if (driver.FindElement(by).Displayed)
                    {
                        return true;
                    }
                    Thread.Sleep(2 * 1000);
                }
                catch (Exception e)
                {
                    Console.WriteLine("waiting element for display...");
                }
            }
            return false;
        }
    }
}
