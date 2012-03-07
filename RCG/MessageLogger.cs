using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace RCG
{
    public class MessageLogger
    {
        public MessageLogger(string logFileName)
        {
            this.logFileName = logFileName;
        }

        private string logFileName = string.Empty;
        public void LogMessage(string message)
        {
            if (!File.Exists(logFileName))
            {
                FileStream f = File.Create(logFileName);
                f.Close();
            }
            Console.WriteLine(message);
            File.AppendAllText(logFileName, DateTime.Now.ToString("yyyy/MM/dd:HH-mm-ss") + " " + message + Environment.NewLine);
        }
    }
}
