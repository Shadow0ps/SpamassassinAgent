using System;
using System.Diagnostics;
using System.Threading;
using System.Globalization;
using System.IO;

namespace SpamassassinAgent
{
    class AgentLogger
    {
        public static short Debug = 4;
        public static short Info = 3;
        public static short Warning = 2;
        public static short Error = 1;
        public static short Fatal = 0;

        private string path;
        private short maxlevel;

        public AgentLogger(string path, short level)
        {
            this.path = path;
            this.maxlevel = level;
        }

        public void log(string message, short level)
        {
            // Don't log it if we dont want it
            if (level > this.maxlevel)
            {
                return;
            }

            // 10 tries 100ms apart
            int tries = 10;
            while (tries > 0)
            {
                try
                {
                    if (!File.Exists(this.path))
                    {
                        File.CreateText(this.path).Close();
                    }
                    StreamWriter log = File.AppendText(this.path);
                    log.WriteLine("[" + DateTime.Now.ToString("g", DateTimeFormatInfo.InvariantInfo) + "] " + Process.GetCurrentProcess().Id.ToString() + " " + message);
                    log.Close();
                    return;
                }
                catch
                {
                    tries -= 1;
                    Thread.Sleep(100);
                }
            }

        }
    }
}
