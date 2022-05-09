using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace CrystalReportsSourceChange
{
    class Program
    {
        private const string ReportDirectory = "\\\\polsage01.aws.polaner.com\\Sage\\Sage 100 Premium\\MAS90\\Reports";
        private static string[] fileEntries;
        private const string DatabaseName = "polsql01.aws.polaner.com";

        static void Main(string[] args)
        {
            GetFileList();
            ProcessFiles();
        }

        private static void ProcessFiles()
        {
            
            foreach (string file in fileEntries)
            {
                ReportDocument report = LoadReport(file);
                foreach(Table table in report.Database.Tables)
                {
                    ChangeServerName(table);
                }
                SaveReport(report, file);
            }
        }

        private static void SaveReport(ReportDocument report, string file)
        {
            report.SaveAs(file);
            report.Dispose();
        }

        private static void ChangeServerName(Table table)
        {
            var logOnInfo = table.LogOnInfo;
            var connectionInfo = logOnInfo.ConnectionInfo;
            connectionInfo.ServerName = DatabaseName;
            connectionInfo.IntegratedSecurity = true;
            table.ApplyLogOnInfo(logOnInfo);
        }

        private static ReportDocument LoadReport(string file)
        {
            ReportDocument report = new ReportDocument();
            report.Load(file);
            return report;
        }

        private static void GetFileList()
        {
            fileEntries = Directory.GetFiles(ReportDirectory, "*.rpt");
        }
    }
}
