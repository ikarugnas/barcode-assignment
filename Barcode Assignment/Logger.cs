using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Barcode_Assignment
{
    public static class Logger
    {
        public static void WriteLog(string logMessage)
        {
            string logPath = @"..\..\log.txt";
            
            using (StreamWriter writer = new StreamWriter(logPath, true))
            {
                writer.WriteLine($"{DateTime.Now} : {logMessage}");
                writer.Close();
            }
        }
    }
}
