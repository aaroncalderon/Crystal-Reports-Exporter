#region Copyright(c) 2003, Teng-Yong Ng
/*
 Copyright(c) 2003, Teng-Yong Ng
  
 This software is provided 'as-is', without any express or implied warranty. In no 
 event will the authors be held liable for any damages arising from the use of this 
 software.
 
  Permission is granted to anyone to use this software for any purpose, including 
 commercial applications, and to alter it and redistribute it freely, subject to the 
 following restrictions:

 1. The origin of this software must not be misrepresented; you must not claim that 
 you wrote the original software. If you use this software in a product, an 
 acknowledgment (see the following) in the product documentation is required.

 Portions Copyright 2003 Teng-Yong Ng

 2. Altered source versions must be plainly marked as such, and must not be 
 misrepresented as being the original software.

 3. This notice may not be removed or altered from any source distribution.
*/

#endregion
/*
file		:crexport.cs
version	:2.0.*
compile	:csc /t:exe /out:crexport.exe /debug+ 
			/r:"H:\Program Files\Common Files\
			Crystal Decisions\1.0\Managed\CrystalDecisions.CrystalReports.Engine.dll" 
			/r:"H:\Program Files\Common Files\
			Crystal Decisions\1.0\Managed\CrystalDecisions.Shared.dll" crexport.cs

Created By: Teng-Yong Ng

Version history
1.1.1152.38285 - First release
1.1.1190.40252 - Bug fixed. Crystal Reports without parameter does not refresh while exporting
1.2.1217.18398 - Allow user to pass multiple values into multiple value parameters 
1.2.1247.16937 - Allow user to pass null value discrete parameter
1.2.1271.40511 - Allow user to specify export html file into single or separated page.
1.2.1419.21168 - zlib/libpng license agreement added.
*/

using System;
using System.Reflection;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;


namespace crexport
{

    #region Exceptions Classes
    public class NullArgumentException : Exception
    {
        // Base Exception class constructors.
        public NullArgumentException()
            : base() { }
        public NullArgumentException(String message)
            : base(message) { }
        public NullArgumentException(String message, Exception innerException)
            : base(message, innerException) { }
    }

    public class InvalidOutputException : Exception
    {
        public InvalidOutputException()
            : base() { }
        public InvalidOutputException(String message)
            : base(message) { }
        public InvalidOutputException(String message, Exception innerException)
            : base(message, innerException) { }
    }

    public class InvalidServerException : Exception
    {
        public InvalidServerException()
            : base() { }
        public InvalidServerException(string message)
            : base(message) { }
        public InvalidServerException(string message, Exception innerException)
            : base(message, innerException) { }
    }

    public class NullParamNameException : Exception
    {
        public NullParamNameException()
            : base() { }
        public NullParamNameException(string message)
            : base(message) { }
        public NullParamNameException(string message, Exception innerException)
            : base(message, innerException) { }
    }

    public class NullParamValueException : Exception
    {
        public NullParamValueException()
            : base() { }
        public NullParamValueException(string message)
            : base(message) { }
        public NullParamValueException(string message, Exception innerException)
            : base(message, innerException) { }
    }

    public class NullExportTypeException : Exception
    {
        public NullExportTypeException()
            : base() { }
        public NullExportTypeException(string message)
            : base(message) { }
        public NullExportTypeException(string message, Exception innerException)
            : base(message, innerException) { }
    }
    #endregion Exceptions Classes

    class crexport
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                ReportInfo rptinfo = new ReportInfo(args);
                if (rptinfo.GetHelp)
                {
                    DisplayMessage(2);
                }
                else
                {

                    ReportDocument Report = new ReportDocument();
                    DisplayMessage(0);
                    try
                    {


                        if (args.Length == 0)
                        {
                            throw new
                                NullArgumentException("No parameter is specified!");
                        }

                        if (rptinfo.ServerName == null)
                        {
                            throw new
                                InvalidServerException("Unspecified Server Name");
                        }

                        if (rptinfo.ReportPath == null || rptinfo.ReportPath == "")
                        {
                            throw new Exception("Invalid Crystal Reports file");
                        }

                        if (rptinfo.OutputPath == null)
                        {
                            throw new
                                InvalidOutputException("Unspecified Output path and filename");
                        }

                        Report.Load(rptinfo.ReportPath, OpenReportMethod.OpenReportByTempCopy);

                        //logon to Database
                        #region Logon to Database

                        TableLogOnInfo logonInfo = new TableLogOnInfo();
                        foreach (Table table in Report.Database.Tables)
                        {
                            logonInfo.ConnectionInfo.ServerName = rptinfo.ServerName;
                            logonInfo.ConnectionInfo.DatabaseName = rptinfo.DatabaseName;
                            if (rptinfo.Username != null)
                            {
                                logonInfo.ConnectionInfo.UserID = rptinfo.Username;
                            }
                            if (rptinfo.Password != null)
                            {
                                logonInfo.ConnectionInfo.Password = rptinfo.Password;
                            }
                            table.ApplyLogOnInfo(logonInfo);
                        }

                        #endregion

                        //Set the export file format
                        #region Setting export file format

                        if (rptinfo.OutputFormat.ToUpper() == "RTF")
                            Report.ExportOptions.ExportFormatType = ExportFormatType.RichText;
                        else if (rptinfo.OutputFormat.ToUpper() == "TXT")
                            Report.ExportOptions.ExportFormatType = ExportFormatType.Text;
                        else if (rptinfo.OutputFormat.ToUpper() == "TAB")
                            Report.ExportOptions.ExportFormatType = ExportFormatType.TabSeperatedText;
                        else if (rptinfo.OutputFormat.ToUpper() == "CSV")
                            Report.ExportOptions.ExportFormatType = ExportFormatType.CharacterSeparatedValues;
                        else if (rptinfo.OutputFormat.ToUpper() == "PDF")
                            Report.ExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                        else if (rptinfo.OutputFormat.ToUpper() == "RPT")
                            Report.ExportOptions.ExportFormatType = ExportFormatType.CrystalReport;
                        else if (rptinfo.OutputFormat.ToUpper() == "DOC")
                            Report.ExportOptions.ExportFormatType = ExportFormatType.WordForWindows;
                        else if (rptinfo.OutputFormat.ToUpper() == "XLS")
                            Report.ExportOptions.ExportFormatType = ExportFormatType.Excel;
                        else if (rptinfo.OutputFormat.ToUpper() == "XLSDATA")
                            Report.ExportOptions.ExportFormatType = ExportFormatType.ExcelRecord;
                        else if (rptinfo.OutputFormat.ToUpper() == "ERTF")
                            Report.ExportOptions.ExportFormatType = ExportFormatType.EditableRTF;
                        else if (rptinfo.OutputFormat.ToUpper() == "XML")
                            Report.ExportOptions.ExportFormatType = ExportFormatType.Xml;
                        else if (rptinfo.OutputFormat.ToUpper() == "HTM")
                        {
                            int i = 1;
                            while (rptinfo.OutputPath.IndexOf(@"\", (rptinfo.OutputPath.Length) - i, i) < 0)
                            {
                                i++;
                            }
                            int lastSlashPos = rptinfo.OutputPath.Length - i + 1; //Last back slash position in OutputPath string
                            string baseFolder = rptinfo.OutputPath.Substring(0, lastSlashPos - 1);

                            HTMLFormatOptions htmlFormatOptions = new HTMLFormatOptions();
                            htmlFormatOptions.HTMLBaseFolderName = baseFolder;
                            htmlFormatOptions.HTMLFileName = rptinfo.OutputPath;
                            htmlFormatOptions.HTMLEnableSeparatedPages = false;
                            htmlFormatOptions.HTMLHasPageNavigator = true;
                            htmlFormatOptions.FirstPageNumber = 1;

                            Report.ExportOptions.ExportFormatType = ExportFormatType.HTML40;
                            Report.ExportOptions.FormatOptions = htmlFormatOptions;
                        }

                        #endregion

                       
                        if (Report.ParameterFields.Count > 0)
                        {
                            ParameterFieldDefinitions paramDefs = Report.DataDefinition.ParameterFields;
                            ParameterValues paramValues = new ParameterValues();
                            List<string> singleParamValue = new List<string>();

                            for (int i = 0; i < paramDefs.Count; i++)
                            {
                                //A report can contains more than One parameters, hence
                                //we loop through all parameters that user has input and match
                                //it with parameter definition of Crystal Reports file.
                                for (int j = 0; j < rptinfo.ReportParameterName.Count; j++)
                                {
                                    if (paramDefs[i].Name == rptinfo.ReportParameterName[j])
                                    {
                                        if (paramDefs[i].EnableAllowMultipleValue && rptinfo.ReportParameterValue[j].IndexOf("|") != -1)
                                        {
                                            singleParamValue = SplitIntoSingleValue(rptinfo.ReportParameterValue[j]); //split multiple value into single value regardless discrete or range

                                            for (int k = 0; k < singleParamValue.Count; k++)
                                                AddParameter(ref paramValues, paramDefs[i].DiscreteOrRangeKind, singleParamValue[k], paramDefs[i].Name);

                                            singleParamValue.Clear();
                                        }
                                        else
                                            AddParameter(ref paramValues, paramDefs[i].DiscreteOrRangeKind, rptinfo.ReportParameterValue[j], paramDefs[i].Name);

                                        paramDefs[i].ApplyCurrentValues(paramValues);
                                        paramValues.Clear();

                                        break; //jump into another user input parameter
                                    }
                                }
                            }
                        }
                        else
                            Report.Refresh();
                            //if report document doesn't come with parameter, just refresh it and job done!
                        
                        Report.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;

                        //DiskFileDestinationOptions
                        DiskFileDestinationOptions diskOptions = new DiskFileDestinationOptions();
                        Report.ExportOptions.DestinationOptions = diskOptions;
                        diskOptions.DiskFileName = rptinfo.OutputPath;

                        Report.Export();
                    }
                    catch (NullArgumentException Er)
                    {
                        Console.WriteLine("\nError: " + Er.Message);
                        DisplayMessage(1);
                    }
                    catch (InvalidOutputException Er)
                    {
                        Console.WriteLine("\nError: " + Er.Message);
                        DisplayMessage(1);
                    }

                    catch (InvalidServerException Er)
                    {
                        Console.WriteLine("\nError: " + Er.Message);
                        DisplayMessage(1);
                    }

                    catch (LogOnException)
                    {
                        Console.WriteLine("\nError: Failed to logon to Database. Check username, password, server name and database name parameter");
                    }

                    catch (LoadSaveReportException)
                    {
                        Console.WriteLine("\nError: Failed to Load or save crystal report");
                    }

                    catch (NullParamNameException Er)
                    {
                        Console.WriteLine("\nError: {0}", Er.Message);
                    }

                    catch (NullExportTypeException Er)
                    {
                        Console.WriteLine("\nError: {0}", Er.Message);
                    }

                    catch (Exception Er)
                    {
                        Console.WriteLine("\nMisc Error: {0}", Er.Message);
                        DisplayMessage(1);
                    }
                    finally
                    {
                        Report.Close();
                    }
                }

            }
            catch (Exception Er)
            {
                Console.WriteLine("\n System Error: {0}", Er.Message);
            }
        }

        public static void DisplayMessage(byte mode)
        {
            Version v = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;

            if (mode == 0)
            {

                Console.WriteLine("\nCrystal Reports Exporter Command Line Utility. Version {0}", v.ToString());
                Console.WriteLine("Copyright(c) 2011 Rainforest Software Solution http://www.rainforestnet.com");
            }
            else if (mode == 1)
            {
                Console.WriteLine("Type \"crexport -?\" for help");
            }
            else if (mode == 2)
            {
                Console.WriteLine("\nCrystal Reports Exporter Command Line Utility. Version {0}", v.ToString());
                Console.WriteLine("Copyright(c) 2011 Rainforest Software Solution http://www.rainforestnet.com");
                Console.WriteLine("crexport Arguments Listing");
                Console.WriteLine("---------------------------------------------------");
                Console.WriteLine("-U database login username");
                Console.WriteLine("-P database login password");
                Console.WriteLine("-F Crystal reports path and filename (Mandatory)");
                Console.WriteLine("-S Database Server Name (instance name)");
                Console.WriteLine("-D Database Name");
                Console.WriteLine("-O Crystal reports Output path and filename");
                Console.WriteLine("-E Export file type.(pdf,doc,xls,rtf,htm,rpt,txt,csv...)");
                Console.WriteLine("-a Parameter value");
                Console.WriteLine("---------------------------------------------------");
                Console.WriteLine("\nExample: C:\\> crexport -U user1 -P mypass -S Server01 -D \"ExtremeDB\" -F c:\\test.rpt -O d:\\test.pdf -a \"Supplier Name:Active Outdoors\" -a \"Date Range:(12-01-2001,12-04-2002)\"");
                Console.WriteLine("Learn more in http://www.rainforestnet.com/crystal-reports-exporter/");
            }
        }

        private static string GetStartValue(string parameterString)
        {
            int delimiter = parameterString.IndexOf(",");
            int leftbracket = parameterString.IndexOf("(");

            if (delimiter == -1 || leftbracket == -1)
                throw new Exception("Invalid Range Parameter value. eg. -a \"parameter name:(1000,2000)\"");

            return parameterString.Substring(leftbracket + 1, delimiter - 1).Trim();
        }

        private static string GetEndValue(string parameterString)
        {
            int delimiter = parameterString.IndexOf(",");
            int rightbracket = parameterString.IndexOf(")");

            if (delimiter == -1 || rightbracket == -1)
                throw new Exception("Invalid Range Parameter value. eg. -a \"parameter name:(1000,2000)\"");

            return parameterString.Substring(delimiter + 1, rightbracket - delimiter - 1).Trim();
        }

        private static List<string> SplitIntoSingleValue(string multipleValueString)
        {
            //if "|" found,means multiple values parameter found
            int pipeStartIndex = 0;
            List<string> singleValue = new List<string>();
            bool loop = true; //loop false when it reaches the last parameter to read

            //pipeIndex is start search position of parameter string
            while (loop)
            {
                if (pipeStartIndex == multipleValueString.LastIndexOf("|") + 1)
                    loop = false;

                if (loop) //if this is not the last parameter
                    singleValue.Add(multipleValueString.Substring(pipeStartIndex, multipleValueString.IndexOf("|", pipeStartIndex + 1) - pipeStartIndex).Trim());
                else
                    singleValue.Add(multipleValueString.Substring(pipeStartIndex, multipleValueString.Length - pipeStartIndex).Trim());

                pipeStartIndex = multipleValueString.IndexOf("|", pipeStartIndex) + 1; //index to the next search of pipe
            }
            return singleValue;
        }

        private static void AddParameter(ref ParameterValues pValues, DiscreteOrRangeKind DoR,string inputString, string pName)
        {
            if (inputString.Trim().Length > 0)
            {
                ParameterValue paraValue;
                if (DoR == DiscreteOrRangeKind.DiscreteValue || (DoR == DiscreteOrRangeKind.DiscreteAndRangeValue && inputString.IndexOf("(") == -1))
                {
                    paraValue = new ParameterDiscreteValue();
                    ((ParameterDiscreteValue)paraValue).Value = inputString;
                    Console.WriteLine("Discrete Parameter : {0} = {1}", pName, ((ParameterDiscreteValue)paraValue).Value);
                    pValues.Add(paraValue);
                    paraValue = null;
                }
                else if (DoR == DiscreteOrRangeKind.RangeValue || (DoR == DiscreteOrRangeKind.DiscreteAndRangeValue && inputString.IndexOf("(") != -1))
                {
                    paraValue = new ParameterRangeValue();
                    ((ParameterRangeValue)paraValue).StartValue = GetStartValue(inputString);
                    ((ParameterRangeValue)paraValue).EndValue = GetEndValue(inputString);
                    Console.WriteLine("Range Parameter : {0} = {1} to {2} ", pName, ((ParameterRangeValue)paraValue).StartValue, ((ParameterRangeValue)paraValue).EndValue);
                    pValues.Add(paraValue);
                    paraValue = null;
                }
            }
        }
    }
}
