using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyLibrary
{
    public class MyClass
    {
        public void MyMethod()
        {
            Console.WriteLine("Hello from MyMethod!");
            MessageBox.Show("调用了新位置", "操作提示");
        }
    }
}