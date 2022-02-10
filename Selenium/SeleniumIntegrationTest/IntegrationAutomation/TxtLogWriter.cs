using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegrationAutomation
{
    public class TxtLogWriter : ILog
    {
        private string _logFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log" + DateTime.Now.ToString("yyyyMM")  + ".txt");
        private string _logResultFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "result" + DateTime.Now.ToString("yyyyMM") + ".txt");
        public void Log(string message)
        {
            var writeLists = new List<string>();
            writeLists.Add(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + message);
            File.AppendAllLines(_logFile, writeLists);
        }

        public void WriteResult(string message)
        {
            var writeLists = new List<string>();
            writeLists.Add(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + message);
            File.AppendAllLines(_logResultFile, writeLists);
        }
    }
}
