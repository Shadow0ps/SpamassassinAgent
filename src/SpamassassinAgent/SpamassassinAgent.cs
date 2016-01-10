namespace SpamassassinAgent
{
    using System;
    using System.IO;
    using System.Diagnostics;
    using System.Reflection;
    using System.Text;
    using System.Linq;
    using System.Collections.Generic;
    using Microsoft.Exchange.Data.Transport;
    using Microsoft.Exchange.Data.Transport.Smtp;
    using System.Globalization;
    using System.Threading;


    public class SpamassassinAgentFactory : SmtpReceiveAgentFactory
    {
        /// <summary>
        /// Directory for storing the data relative to the DLL
        /// </summary>
        private const string RelativeDataPath = @"SpamassassinAgentData\";

        /// <summary>
        /// Configuration filename for GreyList Configuration
        /// </summary>
        private const string ConfigFileName = "SpamassassinConfig.xml";

        /// <summary>
        /// Log filename for logging
        /// </summary>
        private const string LogFile = "SpamassassinLog.txt";

        /// <summary>
        /// GreyList Settings
        /// </summary>
        private SpamassassinSettings spamassassinSettings;

        /// <summary>
        /// Will contain the absolute path for RelativeDataPath
        /// </summary>
        private string dataPath;

        public SpamassassinAgentFactory()
        {
            Assembly currAssembly = Assembly.GetAssembly(this.GetType());
            string assemblyPath = Path.GetDirectoryName(currAssembly.Location);
            this.dataPath = Path.Combine(assemblyPath, RelativeDataPath);

            if (!Directory.Exists(this.dataPath))
            {
                Directory.CreateDirectory(this.dataPath);
            }

            this.spamassassinSettings = new SpamassassinSettings(Path.Combine(this.dataPath, ConfigFileName));
        }

        public override SmtpReceiveAgent CreateAgent(SmtpServer server)
        {
            return new SpamassassinAgent(server, this.spamassassinSettings, this.dataPath, Path.Combine(this.dataPath, LogFile));
        }
    }

    public class SpamassassinAgent : SmtpReceiveAgent
    {

        private SmtpServer server;
        private SpamassassinSettings settings;
        private String dataPath;
        private AgentLogger logger;

        public SpamassassinAgent(SmtpServer server, SpamassassinSettings settings, String dataPath, String logPath)
        {
            this.server = server;
            this.settings = settings;
            this.dataPath = dataPath;
            this.logger = new AgentLogger(logPath, (short)this.settings.LogLevel);
            // Register an OnEndOfData event handler.
            this.OnEndOfData += new EndOfDataEventHandler(this.OnEndOfDataHandler);
        }

        // The OnEndOfDataHandler method is invoked when the entire message has been received.
        public void OnEndOfDataHandler(ReceiveEventSource source, EndOfDataEventArgs eodArgs)
        {
            Byte[] newlineNeedle = Encoding.ASCII.GetBytes("\n");
            Byte[] scoreNeedle = Encoding.ASCII.GetBytes("X-Spam-Score: ");
            Byte[] discardFlag = Encoding.ASCII.GetBytes("X-Spam-Discard: YES\n");
            Byte[] recievedNeedle = Encoding.ASCII.GetBytes("Received: ");
            Byte[] indentedLineNeedle = Encoding.ASCII.GetBytes(" ");
            Double score = 0.0;

            this.logger.log("OnEndOfDataHandler Called", AgentLogger.Debug);
            try
            {
                this.logger.log("OnEndOfDataHandler Info: FROM=" + eodArgs.MailItem.FromAddress.ToString() + ", REMOTE=" + eodArgs.SmtpSession.RemoteEndPoint.Address.ToString(), AgentLogger.Info);
                // Check to make sure SpamAssassin exists at the path it's supposed to
                if (!System.IO.File.Exists(this.settings.SpamassassinPath))
                {
                    this.logger.log("Spamassassin does not exist at path: '" + this.settings.SpamassassinPath + "'. Bypassing.", AgentLogger.Fatal);
                    return;
                }
                // Check to see if this is an internal message
                if (!eodArgs.SmtpSession.IsExternalConnection)
                {
                    foreach (EnvelopeRecipient recipient in eodArgs.MailItem.Recipients)
                    {
                        if (recipient.Address.LocalPart.IndexOf("HealthMailbox") == 0)
                        {
                            this.logger.log("Health Mailbox found in recipients. Bypassing.", AgentLogger.Info);
                            return;
                        }
                    }
                    this.logger.log("Internal Connection Found. Bypassing.", AgentLogger.Info);
                    return;
                }
                // Is the message too big?
                if (eodArgs.MailItem.MimeStreamLength > this.settings.MaxMessageSize)
                {
                    this.logger.log("Message is too large. Increase MaxMessageSize to scan larger messages. MAXSIZE=" + this.settings.MaxMessageSize.ToString() + ", MESSAGESIZE=" + eodArgs.MailItem.MimeStreamLength.ToString(), AgentLogger.Warning);
                    return;
                }
                // Get the message stream
                Stream message = eodArgs.MailItem.GetMimeReadStream();
                List<Byte> messageBytes = new List<Byte>(ReadFully(message));
                byte[] messageByteArray = messageBytes.ToArray();
                message.Close();
                this.logger.log("Message Stream Retrieved. BYTES=" + messageByteArray.Length.ToString(), AgentLogger.Debug);

                // Skip the top number of recieved lines
                int linestart = 0;
                int lineend = 0;
                int linestartmatch = 0;
                for (int i = 0; i < this.settings.SkipRecieved; i++)
                {
                    this.logger.log("Skipping #" + i.ToString() + " Recieved line", AgentLogger.Debug);
                    // Find the first instance of the Recieved header
                    linestart = ByteSearch.Locate(messageByteArray, recievedNeedle, 0);
                    // Loop until we find a line that doesn't start with a space.
                    do
                    {
                        lineend = ByteSearch.Locate(messageByteArray, newlineNeedle, linestart);
                        this.logger.log("Located Line End: " + lineend.ToString(), AgentLogger.Debug);
                        linestartmatch = ByteSearch.Locate(messageByteArray, indentedLineNeedle, lineend);
                        this.logger.log("Located LineStartMatch:" + linestartmatch.ToString(), AgentLogger.Debug);
                        this.logger.log("Bytes between: " + BitConverter.ToString(messageBytes.GetRange(lineend, 5).ToArray()), AgentLogger.Debug);
                        messageBytes.RemoveRange(linestart, lineend - linestart + 1);
                        messageByteArray = messageBytes.ToArray();
                    } while (linestartmatch == lineend + 1);
                    this.logger.log("Finished Skipping #" + i.ToString(), AgentLogger.Debug);

                }

                // Run spamassassin while piping stdin/stdout
                this.logger.log("Starting SpamAssassin Process. PATH='" + this.settings.SpamassassinPath + "', ARGS='" + this.settings.SpamassassinArgs + "'", AgentLogger.Debug);
                Process spamassassin = new Process();
                spamassassin.StartInfo.UseShellExecute = false;
                spamassassin.StartInfo.FileName = this.settings.SpamassassinPath;
                spamassassin.StartInfo.WorkingDirectory = Path.GetDirectoryName(spamassassin.StartInfo.FileName);
                spamassassin.StartInfo.Arguments = ("-s " + this.settings.MaxMessageSize + " " + this.settings.SpamassassinArgs).Trim();
                spamassassin.StartInfo.RedirectStandardInput = true;
                spamassassin.StartInfo.RedirectStandardOutput = true;
                spamassassin.StartInfo.RedirectStandardError = true;
                spamassassin.Start();
                this.logger.log("Started SpamAssassin Process. PID='" + spamassassin.Id.ToString() + "'", AgentLogger.Debug);

                // Copy the message into stdio using a byte for byte copy
                this.logger.log("Copying message to Spamassassin. BYTES=" + messageByteArray.Length.ToString(), AgentLogger.Debug);
                spamassassin.StandardInput.BaseStream.Write(messageByteArray, 0, messageByteArray.Length);
                this.logger.log("Flushing STDIN.", AgentLogger.Debug);
                spamassassin.StandardInput.BaseStream.Flush();
                this.logger.log("Closing STDIN.", AgentLogger.Debug);
                spamassassin.StandardInput.BaseStream.Close();
                this.logger.log("Closed STDIN.", AgentLogger.Debug);

                // Read the entire output buffer, put it into a list for easy manipulation
                this.logger.log("Reading STDOUT from spamassassin.", AgentLogger.Debug);
                List<Byte> outBytes = new List<Byte>(ReadFully(spamassassin.StandardOutput.BaseStream));
                this.logger.log("Read STDOUT. BYTES=" + outBytes.Count.ToString(), AgentLogger.Debug);
                spamassassin.StandardOutput.BaseStream.Close();
                this.logger.log("Closed STDOUT.", AgentLogger.Debug);

                // Read the entire stderr buffer, place it in a file if there's anything to it
                this.logger.log("Reading STDERR from spamassassin.", AgentLogger.Debug);
                Byte[] outErrorBytes = ReadFully(spamassassin.StandardError.BaseStream);
                this.logger.log("Read STDERR. BYTES=" + outErrorBytes.Length.ToString(), AgentLogger.Debug);
                if (outErrorBytes.Length > 0)
                {
                    this.logger.log("Error From SpamAssassin: " + outErrorBytes.ToString(), AgentLogger.Error);
                }
                spamassassin.StandardError.BaseStream.Close();
                this.logger.log("Closed STDERR.", AgentLogger.Debug);

                // Wait for process to exit
                this.logger.log("Waiting for spamassassin to exist.", AgentLogger.Debug);
                spamassassin.WaitForExit();
                this.logger.log("Spamassassin Exited.", AgentLogger.Debug);

                // Find a header


                Int32 flagStart = ByteSearch.Locate(outBytes.ToArray(), scoreNeedle, 0);
                if (flagStart > -1)
                {
                    Int32 scoreStart = flagStart + scoreNeedle.Length;
                    Int32 scoreEnd = ByteSearch.Locate(outBytes.ToArray(), newlineNeedle, scoreStart);
                    if (scoreEnd > -1)
                    {
                        Byte[] scoreBytes = outBytes.GetRange(scoreStart, scoreEnd - scoreStart).ToArray();
                        try
                        {
                            score = Double.Parse(new String(scoreBytes.Select(b => (Char)b).ToArray()).Trim());
                        }
                        catch
                        {

                            // Do nothing
                        }
                        if (score >= this.settings.RejectThreshold)
                        {
                            this.logger.log("Score(" + score.ToString() + ") above threshold(" + this.settings.RejectThreshold.ToString() + "), flagging for discard.", AgentLogger.Info);

                            outBytes.InsertRange(flagStart, discardFlag);
                        }
                        else
                        {
                            this.logger.log("Score(" + score.ToString() + ") below threshold(" + this.settings.RejectThreshold.ToString() + "), passing.", AgentLogger.Info);
                        }
                    }
                    else
                    {
                        this.logger.log("WARNING: Could not find end of score in message.", AgentLogger.Warning);

                    }
                }
                else
                {
                    this.logger.log("WARNING: Could not find start of score in message.", AgentLogger.Warning);
                }

                Byte[] writeBytes = outBytes.ToArray();

                // Now write the data back to a message
                Stream messageOut = eodArgs.MailItem.GetMimeWriteStream();
                this.logger.log("Writing message to MailItem. BYTES=" + writeBytes.Length.ToString(), AgentLogger.Debug);
                messageOut.Write(writeBytes, 0, writeBytes.Length);
                this.logger.log("Closing MailItem buffer.", AgentLogger.Debug);
                messageOut.Close();

            }
            catch (Exception e)
            {
                this.logger.log("Exception Detected", AgentLogger.Fatal);
                this.logger.log(e.ToString(), AgentLogger.Fatal);
            }
            this.logger.log("End of OnEndOfDataHandler", AgentLogger.Debug);
            this.logger.log("----------------------------------------------", AgentLogger.Debug);
            return;
        }
        public static byte[] ReadFully(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }
    }
}
