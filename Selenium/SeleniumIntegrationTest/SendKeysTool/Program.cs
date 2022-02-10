using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SendKeysTool
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string logfile = Path.Combine(AppContext.BaseDirectory, "SendKeysTool.log");
            System.IO.File.WriteAllText(logfile, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + string.Join(",", args));
            foreach(string keyValue in args)
            {
                SendKeys.SendWait(keyValue);
            }
        }
    }
}
