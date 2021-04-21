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
using System.Data.SqlClient;

namespace LimsCrystalReport
{
    public partial class fmMain : Form
    {
        private DBType dbType = DBType.SqlServer;
        public fmMain()
        {
            InitializeComponent();
            string dbTypeConfig = ConfigurationManager.AppSettings["DBType"];
            if (dbTypeConfig == "oracle")
            {
                dbType = DBType.Oracle;
            }
            else
            {
                dbType = DBType.SqlServer;
            }
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
            MessageBox.Show("完成");
        }

        private void btnExportAll_Click(object sender, EventArgs e)
        {
            btnExportAll.Enabled = false;
            //
            string sql = "select count(1) from limsreports " +
                "where categoryid not in " +
                "(select categoryid from limsreportcategories " +
                "where cattype = 'AA')";
            //v10
            sql = "select count(1) from limsreports "  ;
            int count = 1;
            if (dbType == DBType.SqlServer)
            {
                count = (int)SQLHelper.ExecuteScalar(sql);
            }
            else
            { 
                count = int.Parse(OracleHelper.ExecToSqlGetTable(sql).Rows[0][0].ToString());
            }
            pb.Maximum = count;
            //
            bgw.RunWorkerAsync();
        }

        private void ExportAll()
        {
            string exportBasePath = @"D:\Workspace\LimsCrystalReport\Export\";
            exportBasePath = ConfigurationManager.AppSettings["exportBasePath"];
            if (!Directory.Exists(exportBasePath))
            {
                exportBasePath = Environment.CurrentDirectory + @"\Export\";
                Directory.CreateDirectory(exportBasePath);
            }
            //报表分类
            string sql = @"select rc.displaytext, rc.categoryid
                          from limsreportcategories rc where rc.cattype = 'RPT'";
            sql = @"select rc.displaytext, rc.categoryid
                          from limsreportcategories rc  ";
            DataTable dtReportCategory = null;
            if (dbType == DBType.SqlServer)
            {
                dtReportCategory = SQLHelper.GetDataTable(sql);
            }
            else
            {
                dtReportCategory = OracleHelper.ExecToSqlGetTable(sql);
            }
           //循环各分类
            foreach(DataRow drReportCategory in dtReportCategory.Rows)
            {
                string categoryid = drReportCategory["CATEGORYID"].ToString();
                //每个分类创建一个文件夹
                string exportPath = exportBasePath + drReportCategory["DISPLAYTEXT"].ToString() + @"/";
                Directory.CreateDirectory(exportPath);
                //中英文各创建一个子文件夹
                string exportCHSPath = exportPath + @"CHS/";
                Directory.CreateDirectory(exportCHSPath);
                string exportENGPath = exportPath + @"ENG/";
                Directory.CreateDirectory(exportENGPath);
                //查询某分类下报表
                sql = @"select r.reportid, r.displaytext
                    from limsreports r 
                    where r.categoryid = {0}";
                DataTable dtReport = null;
                if (dbType == DBType.SqlServer)
                {
                    dtReport = SQLHelper.GetDataTable(string.Format(sql, "@CATEGORYID")
                    , new SqlParameter[] { new SqlParameter("@CATEGORYID", categoryid) });
                }
                else
                {
                    dtReport = OracleHelper.ExecToSqlGetTable(string.Format(sql, ":CATEGORYID")
                    , new OracleParameter[] { new OracleParameter(":CATEGORYID", categoryid) });
                }
                //循环各报表
                int iReport = 0;
                foreach(DataRow drReport in dtReport.Rows)
                {
                    string reportid = drReport["REPORTID"].ToString();
                    //中、英文报表
                    string exportCHSFile = exportCHSPath+ drReport["DISPLAYTEXT"].ToString() + ".rpt";
                    string exportENGFile = exportENGPath + drReport["DISPLAYTEXT"].ToString() + ".rpt";
                    //查询中英文报表文档
                    sql = @"select rd.report,rd.langid
                        from limsreportdocuments rd 
                        where 1=1 and rd.reportid = {0}";
                    DataTable dtReportDoc = null;
                    if (dbType == DBType.SqlServer)
                    {
                        dtReportDoc = SQLHelper.GetDataTable(string.Format(sql, "@REPORTID")
                        , new SqlParameter[] { new SqlParameter("@REPORTID", reportid) });
                    }
                    else
                    {
                        dtReportDoc = OracleHelper.ExecToSqlGetTable(string.Format(sql, ":REPORTID")
                    , new OracleParameter[] { new OracleParameter(":REPORTID", reportid) });
                    } 
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
