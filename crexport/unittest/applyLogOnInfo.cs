        /// <source>
        /// This code came from
        /// https://archive.sap.com/discussions/thread/1891114
        /// </source>
        /// 
        /// <summary>
        /// Defines the datasource connectivity information for the report
        /// </summary>
        /// <param name="document">Crystal Report ReportDocument</param>
        /// <param name="report">Report object</param>
        private void ApplyLoginInfo(ReportDocument document, Report report)
        {
            TableLogOnInfo info = null;
 
            try
            {
                #region Credentials
 
                //
                // Define credentials
                //
                info = new TableLogOnInfo();
 
                info.ConnectionInfo.AllowCustomConnection = true;                
                info.ConnectionInfo.ServerName = report.Server;
                info.ConnectionInfo.DatabaseName = report.Database;                
 
                //
                // Set the userid/password for the report if we are not using integrated security
                //
                if (report.IntegratedSecurity)
                {
                    info.ConnectionInfo.IntegratedSecurity = true;
                }
                else
                {
                    info.ConnectionInfo.Password = report.Password;
                    info.ConnectionInfo.UserID = report.UserId;
                }
 
                #endregion
 
                #region Apply to connections, tables and sub-reports
 
                //
                // Main connection?
                //
                document.SetDatabaseLogon(info.ConnectionInfo.UserID,
                    info.ConnectionInfo.Password,
                    info.ConnectionInfo.ServerName,
                    info.ConnectionInfo.DatabaseName,
                    false);
 
                //
                // Other connections?
                //
                foreach (CrystalDecisions.Shared.IConnectionInfo connection in document.DataSourceConnections)
                {                    
                    connection.SetConnection(report.Server, report.Database, report.IntegratedSecurity);
                    connection.SetLogon(report.UserId, report.Password);
                    connection.LogonProperties.Set("Data Source", report.Server);
                    connection.LogonProperties.Set("Initial Catalog", report.Database);
                }
 
                //
                // Only do this to the main report (can't do it to sub reports)
                //
                if (!document.IsSubreport)
                {
                    //
                    // Apply to subreports
                    //                
                    foreach (ReportDocument rd in document.Subreports)
                    {
                        ApplyLoginInfo(rd, report);                        
                    }
                }
 
                //
                // Apply to tables
                //
                foreach (CrystalDecisions.CrystalReports.Engine.Table table in document.Database.Tables)
                {
                    TableLogOnInfo tableLogOnInfo = table.LogOnInfo;
                                        
                    tableLogOnInfo.ConnectionInfo = info.ConnectionInfo;                    
                    table.ApplyLogOnInfo(tableLogOnInfo);
                    if (!table.TestConnectivity())
                    {
                        Debug.WriteLine("Failed to apply log in info for Crystal Report");
                    }
                }
 
                #endregion
 
                try
                {
                    //
                    // Break it all down
                    //
                    document.VerifyDatabase();
                }
                catch (LogOnException excLogon)
                {
                    Debug.WriteLine(excLogon.Message);
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Failed to apply login information to the report - " +
                    ex.Message);
            }
        }