namespace SpamassassinAgent
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Interfaces Spamassassin Agent to a settings XML
    /// </summary>
    public class SpamassassinSettings
    {
        /// <summary>
        /// Path to SpamAssassin
        /// </summary>
        private String spamassassinPath;

        /// <summary>
        /// Args for SpamAssassin
        /// </summary>
        private String spamassassinArgs;

        /// <summary>
        /// Threshold for rejection.
        /// </summary>
        private Double rejectThreshold;

        private int logLevel;

        private long maxMessageSize;

        private int skipRecieved;

        /// <summary>
        /// An empty constructor initializes with default values.
        /// </summary>
        /// <param name="path">The path to an XML file that contains the settings.</param>
        public SpamassassinSettings(string path)
        {
            // Default path to use with installed SpamAssassin
            this.spamassassinPath = "C:\\Program Files (x86)\\JAM Software\\SpamAssassin for Windows\\Spamassassin.exe";

            // Default args
            this.spamassassinArgs = "";

            // Default threshold for rejecting
            this.rejectThreshold = 10.0;

            this.logLevel = 2;

            this.maxMessageSize = (1024 * 1024 * 10);

            this.skipRecieved = 0;

            // Read nondefault settings from file.
            this.ReadXMLConfig(path);
        }

        /// <summary>
        /// Clone SpamAsssassinSettings from another object
        /// </summary>
        /// <param name="other">clone from</param>
        public SpamassassinSettings(SpamassassinSettings other)
        {
            this.SpamassassinPath = other.SpamassassinPath;
            this.SpamassassinArgs = other.SpamassassinArgs;
            this.RejectThreshold = other.RejectThreshold;
            this.MaxMessageSize = other.MaxMessageSize;
        }

        /// <summary>
        /// Path to Spamassassin Executable (spamc.exe)
        /// </summary>
        public String SpamassassinPath
        {
            get { return this.spamassassinPath; }

            set { this.spamassassinPath = value; }
        }

        /// <summary>
        /// Additional Arguments to pass to the Spamassassin Executable
        /// </summary>
        public String SpamassassinArgs
        {
            get { return this.spamassassinArgs; }

            set { this.spamassassinArgs = value; }
        }

        /// <summary>
        /// Reject Messages above this threshold
        /// </summary>
        public Double RejectThreshold
        {
            get { return this.rejectThreshold; }

            set { this.rejectThreshold = value; }
        }

        /// <summary>
        /// Maximum size of messages to scan
        /// </summary>
        public long MaxMessageSize
        {
            get { return this.maxMessageSize; }

            set { this.maxMessageSize = value; }
        }

        /// <summary>
        /// Skip these number of Recieved headers
        /// </summary>
        public int SkipRecieved
        {
            get { return this.skipRecieved; }
            set { this.skipRecieved = value; }
        }

        /// <summary>
        /// Maximum log level to log
        /// </summary>
        public int LogLevel
        {
            get { return this.logLevel; }
            set { this.logLevel = value; }
        }
        #region XML File Parsing
        /// <summary>
        /// Reads in configuration options from an XML file and sets the instance
        /// variables to the corresponding values that are read in if they
        /// are valid. If an invalid value is found, or a value is not
        /// set in the XML file, the variable will not be changed from its
        /// default value.
        /// </summary>
        /// <param name="path">The path to the XML configuration file.</param>
        /// <returns>True if the file was read.</returns>
        public void ReadXMLConfig(string path)
        {
            try
            {
                // Some temp variables that will be used during validation.
                String tmpstring = "";
                Double tmpdouble = 0.0;
                int tmpint = 0;
                long tmplong = 0L;
                // Load the file into the XML reader.
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(path);

                XmlNode xmlRoot = xmlDoc.SelectSingleNode("SpamassassinSettings");

                tmpstring = this.ReadXmlString(xmlRoot, "SpamassassinPath");
                if (tmpstring.Length > 0)
                {
                    this.spamassassinPath = tmpstring;
                }

                tmpstring = this.ReadXmlString(xmlRoot, "SpamassassinArgs");
                if (tmpstring.Length > 0)
                {
                    this.spamassassinArgs = tmpstring;
                }

                // Read in the verified entry lifetime.
                tmpdouble = this.ReadXmlDouble(xmlRoot, "RejectThreshold");
                if (tmpdouble > 0.0)
                {
                    this.rejectThreshold = tmpdouble;
                }
                tmpint = this.ReadXmlInt(xmlRoot, "LogLevel");
                if (tmpint != 2)
                {
                    this.logLevel = tmpint;
                }
                tmplong = this.ReadXmlLong(xmlRoot, "MaxMessageSize");
                if (tmpint > 0)
                {
                    this.maxMessageSize = tmplong;
                }
                tmpint = this.ReadXmlInt(xmlRoot, "SkipRecievedHeaders");
                if (tmpint > 0)
                {
                    this.skipRecieved = tmpint;
                }
            }
            catch (XmlException e)
            {
                Debug.WriteLine(e.ToString());
                return;
            }
            catch (UnauthorizedAccessException e)
            {
                Debug.WriteLine(e.ToString());
                return;
            }
            catch (IOException e)
            {
                Debug.WriteLine(e.ToString());
                return;
            }
            return;
        }

        private int ReadXmlInt(XmlNode root, string xmlParam)
        {
            int retval = 0;

            if (root != null && xmlParam != null)
            {
                XmlNode valNode = root.SelectSingleNode(xmlParam);
                if (valNode != null)
                {
                    XmlNode childNode = valNode.FirstChild;
                    if (childNode != null)
                    {
                        int.TryParse(childNode.Value, out retval);
                    }
                }
            }

            return retval;
        }

        private long ReadXmlLong(XmlNode root, string xmlParam)
        {
            long retval = 0;

            if (root != null && xmlParam != null)
            {
                XmlNode valNode = root.SelectSingleNode(xmlParam);
                if (valNode != null)
                {
                    XmlNode childNode = valNode.FirstChild;
                    if (childNode != null)
                    {
                        long.TryParse(childNode.Value, out retval);
                    }
                }
            }

            return retval;
        }

        private double ReadXmlDouble(XmlNode root, string xmlParam)
        {
            Double retval = 0;

            if (root != null && xmlParam != null)
            {
                XmlNode valNode = root.SelectSingleNode(xmlParam);
                if (valNode != null)
                {
                    XmlNode childNode = valNode.FirstChild;
                    if (childNode != null)
                    {
                        Double.TryParse(childNode.Value, out retval);
                    }
                }
            }

            return retval;
        }

        private String ReadXmlString(XmlNode root, string xmlParam)
        {
            String retval = "";

            if (root != null && xmlParam != null)
            {
                XmlNode xmlParamNode = root.SelectSingleNode(xmlParam);
                if (xmlParamNode != null)
                {
                    XmlNode childNode = xmlParamNode.FirstChild;
                    if (childNode != null)
                    {
                        retval = childNode.Value.ToString();
                    }
                }
            }

            return retval;
        }
        #endregion XML File Parsing
    }
}
