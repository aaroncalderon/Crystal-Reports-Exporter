namespace applyLogOnInfo
{
    class aplylogoninfo
    {
        /// <source>
        /// This code came from
        /// https://archive.sap.com/discussions/thread/1891114
        /// </source>
        /// 
        /// <summary>
        /// Defines the datasource connectivity information for the rptinfo
        /// </summary>
        /// <param name="Report">Crystal Report ReportDocument</param>
        /// <param name="rptinfo">Report object</param>
        private void ApplyLoginInfo(ReportDocument Report, Report rptinfo)
        {
            TableLogOnInfo logonInfo = null;

            try
            {
                #region Credentials

                //
                // Define credentials
                //
                logonInfo = new TableLogOnInfo();

                logonInfo.ConnectionInfo.AllowCustomConnection = true;
                logonInfo.ConnectionInfo.ServerName = rptinfo.Server;
                logonInfo.ConnectionInfo.DatabaseName = rptinfo.Database;

                //
                // Set the userid/password for the rptinfo if we are not using integrated security
                //
                if (rptinfo.IntegratedSecurity)
                {
                    logonInfo.ConnectionInfo.IntegratedSecurity = true;
                }
                else
                {
                    logonInfo.ConnectionInfo.Password = rptinfo.Password;
                    logonInfo.ConnectionInfo.UserID = rptinfo.UserId;
                }

                #endregion

                #region Apply to connections, tables and sub-reports

                //
                // Main connection?
                //
                Report.SetDatabaseLogon(logonInfo.ConnectionInfo.UserID,
                    logonInfo.ConnectionInfo.Password,
                    logonInfo.ConnectionInfo.ServerName,
                    logonInfo.ConnectionInfo.DatabaseName,
                    false);

                //
                // Other connections?
                //
                foreach (CrystalDecisions.Shared.IConnectionInfo connection in Report.DataSourceConnections)
                {
                    connection.SetConnection(rptinfo.Server, rptinfo.Database, rptinfo.IntegratedSecurity);
                    connection.SetLogon(rptinfo.UserId, rptinfo.Password);
                    connection.LogonProperties.Set("Data Source", rptinfo.Server);
                    connection.LogonProperties.Set("Initial Catalog", rptinfo.Database);
                }

                //
                // Only do this to the main rptinfo (can't do it to sub reports)
                //
                if (!Report.IsSubreport)
                {
                    //
                    // Apply to subreports
                    //                
                    foreach (ReportDocument rd in Report.Subreports)
                    {
                        ApplyLoginInfo(rd, rptinfo);
                    }
                }

                //
                // Apply to tables
                //
                foreach (CrystalDecisions.CrystalReports.Engine.Table table in Report.Database.Tables)
                {
                    TableLogOnInfo tableLogOnInfo = table.LogOnInfo;

                    tableLogOnInfo.ConnectionInfo = logonInfo.ConnectionInfo;
                    table.ApplyLogOnInfo(tableLogOnInfo);
                    if (!table.TestConnectivity())
                    {
                        Debug.WriteLine("Failed to apply log in logonInfo for Crystal Report");
                    }
                }

                #endregion

                try
                {
                    //
                    // Break it all down
                    //
                    Report.VerifyDatabase();
                }
                catch (LogOnException excLogon)
                {
                    Debug.WriteLine(excLogon.Message);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Failed to apply login information to the rptinfo - " +
                    ex.Message);
            }
        }
    }
}