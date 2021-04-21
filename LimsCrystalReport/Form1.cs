using LimsCrystalReport.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Oracle.ManagedDataAccess.Client;
using System.Configuration;

namespace LimsCrystalReport
{
    public partial class fmMain : Form
    {
        public fmMain()
        {
            InitializeComponent();
        }

        private void bgw_DoWork(object sender, DoWorkEventArgs e)
        {
            ExportAll();
        }

        private void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pb.Value++;
        }

        private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnExportAll.Enabled = true;
        }

        private void btnExportAll_Click(object sender, EventArgs e)
        {
            btnExportAll.Enabled = false;
            //
            string sql = "select count(1) from limsreports " +
                "where categoryid not in " +
                "(select categoryid from limsreportcategories " +
                "where cattype = 'AA')";
            int count = 1;
            //count = (int)SQLHelper.ExecuteScalar(sql);
            count = int.Parse(OracleHelper.ExecToSqlGetTable(sql).Rows[0][0].ToString());
            pb.Maximum = count;
            //
            bgw.RunWorkerAsync();
        }

        private void ExportAll()
        {
            string exportBasePath = @"D:\Workspace\LimsCrystalReport\Export\";
            exportBasePath = ConfigurationManager.AppSettings["exportBasePath"];
            //
            string sql = @"select rc.displaytext, rc.categoryid
                          from limsreportcategories rc where rc.cattype = 'RPT'";
            DataTable dtReportCategory = OracleHelper.ExecToSqlGetTable(sql);
            foreach(DataRow drReportCategory in dtReportCategory.Rows)
            {
                string categoryid = drReportCategory["CATEGORYID"].ToString();
                //
                string exportPath = exportBasePath + drReportCategory["DISPLAYTEXT"].ToString() + @"/";
                Directory.CreateDirectory(exportPath);
                string exportCHSPath = exportPath + @"CHS/";
                Directory.CreateDirectory(exportCHSPath);
                string exportENGPath = exportPath + @"ENG/";
                Directory.CreateDirectory(exportENGPath);
                //
                sql = @"select r.reportid, r.displaytext
                    from limsreports r 
                    where r.categoryid = {0}";
                DataTable dtReport = OracleHelper.ExecToSqlGetTable(string.Format(sql, ":CATEGORYID")
                    , new OracleParameter[] { new OracleParameter(":CATEGORYID", categoryid) });
                int iReport = 0;
                foreach(DataRow drReport in dtReport.Rows)
                {
                    string reportid = drReport["REPORTID"].ToString();
                    //
                    string exportCHSFile = exportCHSPath+ drReport["DISPLAYTEXT"].ToString() + ".rpt";
                    string exportENGFile = exportENGPath + drReport["DISPLAYTEXT"].ToString() + ".rpt";
                    //
                    sql = @"select rd.report,rd.langid
                        from limsreportdocuments rd 
                        where 1=1 and rd.reportid = {0}";
                    DataTable dtReportDoc = OracleHelper.ExecToSqlGetTable(string.Format(sql, ":REPORTID")
                    , new OracleParameter[] { new OracleParameter(":REPORTID", reportid) });
                    foreach(DataRow drReportDoc in dtReportDoc.Rows)
                    {
                        string langid = drReportDoc["LANGID"].ToString();
                        byte[] reportDocContent = (byte[])drReportDoc["REPORT"];
                        if (langid == "CHS")
                        {
                            File.WriteAllBytes(exportCHSFile, reportDocContent);
                        }else if(langid == "ENG")
                        {
                            File.WriteAllBytes(exportENGFile, reportDocContent);
                        }
                    }
                    bgw.ReportProgress(++iReport);
                }
            }
        }
    }
}
