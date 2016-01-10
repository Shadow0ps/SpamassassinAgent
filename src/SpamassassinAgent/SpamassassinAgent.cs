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


    /// <summary>
    /// Agent Factory for SpamAssassin
    /// </summary>
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

        /// <summary>
        /// Agent Factory
        /// </summary>
        public SpamassassinAgentFactory()
        {
            // Get the current location of where this agent is executing from
            Assembly currAssembly = Assembly.GetAssembly(this.GetType());
            string assemblyPath = Path.GetDirectoryName(currAssembly.Location);
            this.dataPath = Path.Combine(assemblyPath, RelativeDataPath);

            // If the data directory doesn't exist ...
            if (!Directory.Exists(this.dataPath))
            {
                // ... Create it
                Directory.CreateDirectory(this.dataPath);
            }
            // Fetch SpamassassinAgent settings
            this.spamassassinSettings = new SpamassassinSettings(Path.Combine(this.dataPath, ConfigFileName));
        }

        /// <summary>
        /// Spawn an SMTP recieve agent
        /// </summary>
        /// <param name="server">Exchange SmtpServer resource</param>
        /// <returns>Spawned Agent</returns>
        public override SmtpReceiveAgent CreateAgent(SmtpServer server)
        {
            return new SpamassassinAgent(this.spamassassinSettings, this.dataPath, Path.Combine(this.dataPath, LogFile));
        }
    }

    /// <summary>
    /// Main class for SpamassassinAgent
    /// </summary>
    public class SpamassassinAgent : SmtpReceiveAgent
    {

        /// <summary>
        /// Populated Spamassassin Settings
        /// </summary>
        private SpamassassinSettings settings;

        /// <summary>
        /// Absolute data path for storing the agent data
        /// </summary>
        private String dataPath;

        /// <summary>
        /// Logging object
        /// </summary>
        private AgentLogger logger;

        /// <summary>
        /// Initializer for SpamassassinAgent
        /// </summary>
        /// <param name="settings"></param>
        /// <param name="dataPath"></param>
        /// <param name="logPath"></param>
        public SpamassassinAgent(SpamassassinSettings settings, String dataPath, String logPath)
        {
            this.settings = settings;
            this.dataPath = dataPath;

            // Spawn a new logger
            this.logger = new AgentLogger(logPath, (short)this.settings.LogLevel);

            // Register an OnEndOfData event handler.
            this.OnEndOfData += new EndOfDataEventHandler(this.OnEndOfDataHandler);
        }

        /// <summary>
        /// Handles processing the message once the end of the DATA SMTP stream is sent.
        /// </summary>
        /// <param name="source"></param>
        /// <param name="eodArgs"></param>
        public void OnEndOfDataHandler(ReceiveEventSource source, EndOfDataEventArgs eodArgs)
        {
            Byte[] newlineNeedle = Encoding.ASCII.GetBytes("\n");
            Byte[] scoreNeedle = Encoding.ASCII.GetBytes("X-Spam-Score: ");
            Byte[] discardFlag = Encoding.ASCII.GetBytes("X-Spam-Discard: YES\n");
            Byte[] recievedNeedle = Encoding.ASCII.GetBytes("Received: ");
            Byte[] indentedLineNeedle = Encoding.ASCII.GetBytes(" ");
            Nullable<Double> score = null;

            this.logger.debug("OnEndOfDataHandler Called");

            // Wrap everything in a Try. This makes sure that even if something fails the message still passes through
            try
            {
                this.logger.info("OnEndOfDataHandler Info: FROM=" + eodArgs.MailItem.FromAddress.ToString() + ", REMOTE=" + eodArgs.SmtpSession.RemoteEndPoint.Address.ToString());

                // Check to make sure SpamAssassin exists at the path it's supposed to
                if (!System.IO.File.Exists(this.settings.SpamassassinPath))
                {
                    this.logger.fatal("Spamassassin does not exist at path: '" + this.settings.SpamassassinPath + "'. Bypassing.");
                    this.logger.flush();
                    return;
                }

                // Check to see if this is an internal message
                if (!eodArgs.SmtpSession.IsExternalConnection)
                {
                    foreach (EnvelopeRecipient recipient in eodArgs.MailItem.Recipients)
                    {
                        if (recipient.Address.LocalPart.IndexOf("HealthMailbox") == 0)
                        {
                            this.logger.info("Health Mailbox found in recipients. Bypassing.");
                            this.logger.flush();
                        }
                    }
                    this.logger.info("Internal Connection Found. Bypassing.");
                    this.logger.flush();
                }

                // Is the message too big?
                if (eodArgs.MailItem.MimeStreamLength > this.settings.MaxMessageSize)
                {
                    this.logger.warning("Message is too large. Increase MaxMessageSize to scan larger messages. MAXSIZE=" + this.settings.MaxMessageSize.ToString() + ", MESSAGESIZE=" + eodArgs.MailItem.MimeStreamLength.ToString());
                    this.logger.flush();
                }

                // Get the message stream
                Stream message = eodArgs.MailItem.GetMimeReadStream();
                List<Byte> messageBytes = new List<Byte>(ReadFully(message));
                byte[] messageByteArray = messageBytes.ToArray();
                message.Close();
                this.logger.debug("Message Stream Retrieved. BYTES=" + messageByteArray.Length.ToString());

                // Skip the top number of recieved lines
                int linestart = 0;
                int lineend = 0;
                int linestartmatch = 0;
                for (int i = 0; i < this.settings.SkipRecieved; i++)
                {
                    this.logger.debug("Skipping #" + i.ToString() + " Recieved line");
                    // Find the first instance of the Recieved header
                    linestart = ByteSearch.Locate(messageByteArray, recievedNeedle, 0);
                    // Loop until we find a line that doesn't start with a space.
                    do
                    {
                        lineend = ByteSearch.Locate(messageByteArray, newlineNeedle, linestart);
                        this.logger.debug("Located Line End: " + lineend.ToString());
                        linestartmatch = ByteSearch.Locate(messageByteArray, indentedLineNeedle, lineend);
                        this.logger.debug("Located LineStartMatch:" + linestartmatch.ToString());
                        messageBytes.RemoveRange(linestart, lineend - linestart + 1);
                        messageByteArray = messageBytes.ToArray();

                    } while (linestartmatch == lineend + 1);

                    this.logger.debug("Finished Skipping #" + i.ToString());

                }

                // Run spamassassin while piping stdin/stdout
                this.logger.debug("Starting SpamAssassin Process. PATH='" + this.settings.SpamassassinPath + "', ARGS='" + this.settings.SpamassassinArgs + "'");
                Process spamassassin = new Process();
                spamassassin.StartInfo.UseShellExecute = false;
                spamassassin.StartInfo.FileName = this.settings.SpamassassinPath;
                spamassassin.StartInfo.WorkingDirectory = Path.GetDirectoryName(spamassassin.StartInfo.FileName);
                spamassassin.StartInfo.Arguments = ("-s " + this.settings.MaxMessageSize + " " + this.settings.SpamassassinArgs).Trim();
                spamassassin.StartInfo.RedirectStandardInput = true;
                spamassassin.StartInfo.RedirectStandardOutput = true;
                spamassassin.StartInfo.RedirectStandardError = true;
                spamassassin.Start();
                this.logger.debug("Started SpamAssassin Process. PID='" + spamassassin.Id.ToString() + "'");

                // Copy the message into stdio using a byte for byte copy
                this.logger.debug("Copying message to Spamassassin. BYTES=" + messageByteArray.Length.ToString());
                spamassassin.StandardInput.BaseStream.Write(messageByteArray, 0, messageByteArray.Length);
                spamassassin.StandardInput.BaseStream.Flush();
                spamassassin.StandardInput.BaseStream.Close();

                // Read the entire output buffer, put it into a list for easy manipulation
                List<Byte> outBytes = new List<Byte>(ReadFully(spamassassin.StandardOutput.BaseStream));
                this.logger.debug("Read STDOUT from spamassassin. BYTES=" + outBytes.Count.ToString());
                spamassassin.StandardOutput.BaseStream.Close();

                // Read the entire stderr buffer, place it in a file if there's anything to it
                Byte[] outErrorBytes = ReadFully(spamassassin.StandardError.BaseStream);
                this.logger.debug("Read STDERR from spamassassin. BYTES=" + outErrorBytes.Length.ToString());
                if (outErrorBytes.Length > 0)
                {
                    this.logger.error("Error From SpamAssassin: " + outErrorBytes.ToString());
                }
                spamassassin.StandardError.BaseStream.Close();

                // Wait for process to exit
                spamassassin.WaitForExit();

                // Find a header
                int flagStart = ByteSearch.Locate(outBytes.ToArray(), scoreNeedle, 0);
                int scoreEnd = -1;
                int scoreStart = 0;

                // Check if we found the start of the header
                if (flagStart > -1)
                {
                    scoreStart = flagStart + scoreNeedle.Length;
                    scoreEnd = ByteSearch.Locate(outBytes.ToArray(), newlineNeedle, scoreStart);
                } else {
                    this.logger.warning("Could not find start of score in message.");
                }

                // Check if we found the end of the header
                if (scoreEnd > -1)
                {
                    Byte[] scoreBytes = outBytes.GetRange(scoreStart, scoreEnd - scoreStart).ToArray();
                    try
                    {
                        score = Double.Parse(new String(scoreBytes.Select(b => (Char)b).ToArray()).Trim());
                    }
                    catch(Exception e)
                    {
                        this.logger.warning("Could not parse score. Exception:" + e.Message);
                    }
                }
                else
                {
                    this.logger.warning("Could not find end of score in message.");
                }

                // If we found a score check it and process it
                if(score != null) {
                    if (score >= this.settings.RejectThreshold)
                    {
                        this.logger.info("Score(" + score.ToString() + ") above threshold(" + this.settings.RejectThreshold.ToString() + "), flagging for discard.");

                        outBytes.InsertRange(flagStart, discardFlag);
                    }
                    else
                    {
                        this.logger.info("Score(" + score.ToString() + ") below threshold(" + this.settings.RejectThreshold.ToString() + "), passing.");
                    }
                }

                // Write the message back to SpamAssassin
                Byte[] writeBytes = outBytes.ToArray();
                Stream messageOut = eodArgs.MailItem.GetMimeWriteStream();
                this.logger.debug("Writing message to MailItem. BYTES=" + writeBytes.Length.ToString());
                messageOut.Write(writeBytes, 0, writeBytes.Length);
                messageOut.Close();

            }
            catch (Exception e)
            {
                this.logger.fatal("Exception Detected");
                this.logger.fatal(e.ToString());
            }
            finally
            {
                this.logger.flush();
            }
            
        }

        /// <summary>
        /// Reads a Stream completely into a byte array
        /// </summary>
        /// <param name="input">Stream to read from</param>
        /// <returns></returns>
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
