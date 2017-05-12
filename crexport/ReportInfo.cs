using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace crexport
{
    class ReportInfo
    {
        private string username;                        //-U Report database login username (mandatory)
        private string password;                        //-P Report database login password (mandatory)
        private string reportPath;                      //-F Crystal Report path and filename (mandatory)
        private string outputPath;                      //-O Output file path and filename 
        private string serverName;                      //-S Server name of specified crystal Report (mandatory)
        private string databaseName;                    //-D Database name of specified crystal Report
        private string outputFormat;                    //-E Crystal Report exported format (pdf,xls,htm)
        private List<string> reportParameterString;     //-p Report Parameter set (small letter p). eg:{customer : "Microsoft Corporation"} or {customer : "Microsoft Corporation" | "Google Inc"}

        private List<string> reportParameterName;       //name of the parameter
        private List<string> reportParameterValue;      //the value might not be actual value to pass to ReportDocument as it can be multi value or range value
        private bool getHelp; 

        public ReportInfo(string[] parameters)
        {
            string defaultFileFormat = "txt";
            getHelp = false;

            reportParameterString = new List<string>();
            reportParameterName = new List<string>();
            reportParameterValue = new List<string>();

            #region Assigning crexport parameters to variables
            for (int i = 0; i < parameters.Count(); i++)
            {
                if (i + 1 < parameters.Count())
                {
                    if (parameters[i + 1].Length > 0)
                    {
                        if (parameters[i] == "-U")
                            username = parameters[i + 1];
                        else if (parameters[i] == "-P")
                            password = parameters[i + 1];
                        else if (parameters[i] == "-F")
                            reportPath = parameters[i + 1];
                        else if (parameters[i] == "-O")
                            outputPath = parameters[i + 1];
                        else if (parameters[i] == "-S")
                            serverName = parameters[i + 1];
                        else if (parameters[i] == "-D")
                            databaseName = parameters[i + 1];
                        else if (parameters[i] == "-E")
                            outputFormat = parameters[i + 1];
                        else if (parameters[i] == "-a")
                            reportParameterString.Add(parameters[i + 1]);
                    }
                }
                
                if (parameters[i] == "-?" || parameters[i] == "/?")
                    getHelp = true;
            }
            #endregion

            foreach (string input in reportParameterString)
            {
                ProcessParameter(input);
            }

            #region handle default output path and output format

            if (reportPath != null && reportPath.Length > 0)
            {
                string fileExt = "";                                //default file extension if output filename and path isn't specified.
                if (outputPath == null || outputPath.Trim() == "")  //if output path and filename is not specified
                {
                    if (outputFormat == null)
                        outputFormat = defaultFileFormat;

                    if (outputFormat == "xlsdata")
                        fileExt = "xls";
                    else if (outputFormat == "tab")
                        fileExt = "txt";
                    else if (outputFormat == "ertf")
                        fileExt = "rtf";
                    else
                        fileExt = outputFormat;

                    if (reportPath.LastIndexOf(".rpt") == -1)
                        throw new Exception("Invalid Crystal Reports file");

                    outputPath = String.Format("{0}-{1}.{2}", reportPath.Substring(0, reportPath.LastIndexOf(".rpt")), DateTime.Now.ToString("yyyyMMddHHmmss"), fileExt);
                }
                else
                {
                    //if output path and filename is specified, use file extension to determine output format.
                    if (outputFormat == null)
                    {
                        int lastIndexDot = outputPath.LastIndexOf(".");
                        fileExt = outputPath.Substring(lastIndexDot + 1, 3);

                        //ensure filename extension has 3 char after the dot (.)
                        if ((outputPath.Length == lastIndexDot + 4) && (fileExt == "rtf" || fileExt == "txt" || fileExt == "csv" || fileExt == "pdf" || fileExt == "rpt" || fileExt == "doc" || fileExt == "xls" || fileExt == "xml" || fileExt == "htm"))
                            outputFormat = outputPath.Substring(lastIndexDot + 1, 3);
                        else
                            outputFormat = defaultFileFormat;
                    }
                }
            }
            else
            {
                if (!getHelp)
                    throw new Exception("Invalid Crystal Reports file");
            }

            #endregion
        }

        private void ProcessParameter(string paraString)
        {
            reportParameterName.Add(paraString.Substring(0, paraString.IndexOf(":")).Trim());
            reportParameterValue.Add((paraString.Substring(paraString.IndexOf(":") + 1, paraString.Length - (paraString.IndexOf(":") + 1))).Trim());
        }

        #region Public Properties

        public bool GetHelp
        {
            get
            {
                return getHelp;
            }
        }

        public string Username
        {
            get
            {
                return username;
            }

        }
        public string Password
        {
            get
            {
                return password;
            }
        }

        public string ReportPath
        {
            get
            {
                return reportPath;
            }
        }
        public string OutputPath
        {
            get
            {
                return outputPath;
            }
        }

        public string ServerName
        {
            get
            {
                return serverName;
            }
        }

        public string DatabaseName
        {
            get
            {
                return databaseName;
            }
        }

        public string OutputFormat
        {
            get
            {
                return outputFormat;
            }
        }

        public List<string> ReportParameterName
        {
            get
            {
                return reportParameterName;
            }
        }

        public List<string> ReportParameterValue
        {
            get
            {
                return reportParameterValue;
            }
        }

        #endregion

    }
}
