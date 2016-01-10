using System;
using System.Diagnostics;
using System.Threading;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;

namespace SpamassassinAgent
{
    /// <summary>
    /// Simple Logging class for Exchange Transport Agents
    /// </summary>
    class AgentLogger
    {
        // Various Log levels
        public static short Debug = 4;
        public static short Info = 3;
        public static short Warning = 2;
        public static short Error = 1;
        public static short Fatal = 0;

        /// <summary>
        /// Path to save the log file to
        /// </summary>
        private string path;

        /// <summary>
        /// Maximum level we want to log anything above this gets discarded
        /// </summary>
        private short maxlevel;

        /// <summary>
        /// Buffer for storing logs
        /// </summary>
        private string buf;

        /// <summary>
        /// New AgentLogger
        /// </summary>
        /// <param name="path">Path to save log file to</param>
        /// <param name="level">Maximum log level to save</param>
        public AgentLogger(string path, short level)
        {
            this.path = path;
            this.maxlevel = level;
            this.buf = "";
        }

        /// <summary>
        /// Shortcut for log(message, AgentLogger.Debug)
        /// </summary>
        /// <param name="message"></param>
        /// <param name="callingFilePath"></param>
        /// <param name="callingFileLineNumber"></param>
        public void debug(string message, [CallerFilePath] string callingFilePath = "", [CallerLineNumber] int callingFileLineNumber = 0)
        {
            this.log(message, Debug, callingFilePath, callingFileLineNumber);
        }

        /// <summary>
        /// Shortcut for log(message, AgentLogger.Info)
        /// </summary>
        /// <param name="message"></param>
        /// <param name="callingFilePath"></param>
        /// <param name="callingFileLineNumber"></param>
        public void info(string message, [CallerFilePath] string callingFilePath = "", [CallerLineNumber] int callingFileLineNumber = 0)
        {
            this.log(message, Info, callingFilePath, callingFileLineNumber);
        }

        /// <summary>
        /// Shortcut for log(message, AgentLogger.Warning)
        /// </summary>
        /// <param name="message"></param>
        /// <param name="callingFilePath"></param>
        /// <param name="callingFileLineNumber"></param>
        public void warning(string message, [CallerFilePath] string callingFilePath = "", [CallerLineNumber] int callingFileLineNumber = 0)
        {
            this.log(message, Warning, callingFilePath, callingFileLineNumber);
        }

        /// <summary>
        /// Shortcut for log(message, AgentLogger.Error)
        /// </summary>
        /// <param name="message"></param>
        /// <param name="callingFilePath"></param>
        /// <param name="callingFileLineNumber"></param>
        public void error(string message, [CallerFilePath] string callingFilePath = "", [CallerLineNumber] int callingFileLineNumber = 0)
        {
            this.log(message, Error, callingFilePath, callingFileLineNumber);
        }

        /// <summary>
        /// Shortcut for log(message, AgentLogger.Fatal)
        /// </summary>
        /// <param name="message"></param>
        /// <param name="callingFilePath"></param>
        /// <param name="callingFileLineNumber"></param>
        public void fatal(string message, [CallerFilePath] string callingFilePath = "", [CallerLineNumber] int callingFileLineNumber = 0)
        {
            this.log(message, Fatal, callingFilePath, callingFileLineNumber);
        }

        /// <summary>
        /// Shortcut for log(message, AgentLogger.debug)
        /// </summary>
        /// <param name="message">Message to log</param>
        /// <param name="level">Level for this log line</param>
        /// <param name="callingFilePath">Calling method file (populated by shortcuts)</param>
        /// <param name="callingFileLineNumber">Calling file line (populated by shortcuts)</param>
        public void log(string message, short level, string callingFilePath, int callingFileLineNumber)
        {
            // Don't log it if we dont want it
            if (level > this.maxlevel)
            {
                return;
            }

            this.buf += "[" + DateTime.Now.ToString("g", DateTimeFormatInfo.InvariantInfo) + "] " + callingFilePath + ":" + callingFileLineNumber + " " + message + "\n";
        }

        /// <summary>
        /// Flush the buffer to a file and then clear the buffer
        /// </summary>
        public void flush()
        {
            // If there's nothing in the buffer there's no reason to log
            if (this.buf.Length < 1)
            {
                return;
            }
            
            // Append a separator 
            this.buf += "----------------------------------------------\n";

            // 10 tries 100ms apart just in case the file is locked
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
                    log.Write(buf);
                    log.Close();
                    return;
                }
                catch
                {
                    tries -= 1;
                    // Sleep for 100ms, should be enough time for something elese to log and close.
                    Thread.Sleep(100);
                }
            }

            // Empty the buffer
            this.buf = "";
        }
    }
}
