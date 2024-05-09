using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject2.TestData
{
    public class DataRegister
    {
        public string _FullName { get; set; }
        public string _email { get; set; }
        public string _passWord { get; set; }

        public string _repassword { get; set; }

        public string _expectedResult { get; set; }
        public int row { get; set; }
        public int col { get; set; }
    }
}
