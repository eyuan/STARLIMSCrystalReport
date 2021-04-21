using System;
using System.Collections.Generic; 
using System.Text;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;

namespace LimsCrystalReport.Utils
{
    public static class SQLHelper
    {
        //读取配置文件中的连接字符串
        static string connstr = ConfigurationManager.ConnectionStrings["sql"].ConnectionString;
        /// <summary>
        /// 执行非查询sql语句，返回受影响行数，如果执行非增删改则返回-1
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <param name="paras">参数数组</param>
        /// <returns>影响行数res</returns>
        public static int ExecuteNonQuery(string sql, params SqlParameter[] paras)
        {
            int res = -1;
            using (SqlConnection conn = new SqlConnection(connstr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    //cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandType = CommandType.Text;
                    if (paras != null || paras.Length > 0)
                    {
                        cmd.Parameters.AddRange(paras);
                    }
                    conn.Open();
                    res = cmd.ExecuteNonQuery();
                }
            }
            return res;
        }
        /// <summary>
        /// 执行非查询sql语句，返回受影响行数，如果执行非增删改则返回-1
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <param name="paras">参数数组</param>
        /// <returns>影响行数res</returns>
        public static int ExecuteNonParaQuery(string sql)
        {
            int res = -1;
            using (SqlConnection conn = new SqlConnection(connstr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    conn.Open();
                    res = cmd.ExecuteNonQuery();
                }
            }
            return res;
        }
        /// <summary>
        /// 执行读取数据，返回一个对象
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <param name="paras">参数数组</param>
        /// <returns>返回一个对象o</returns>
        public static object ExecuteScalar(string sql, params SqlParameter[] paras)
        {
            object o = null;
            using (SqlConnection conn = new SqlConnection(connstr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    //cmd.CommandType = CommandType.StoredProcedure;
                    if (paras != null)
                    {
                        cmd.Parameters.AddRange(paras);
                    }
                    conn.Open();
                    o = cmd.ExecuteScalar();
                }
            }
            return o;
        }
        /// <summary>
        /// 执行查询sql语句，返回一个对象
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <param name="paras">查询参数</param>
        /// <returns>返回DataReader对象</returns>
        public static SqlDataReader ExecuteReader(string sql, params SqlParameter[] paras)
        {
            SqlConnection conn = new SqlConnection(connstr);
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.CommandType = CommandType.StoredProcedure;
                if (paras != null)
                {
                    cmd.Parameters.AddRange(paras);
                }
                conn.Open();
                try
                {
                    return cmd.ExecuteReader(CommandBehavior.CloseConnection);
                }
                catch (Exception ex)
                {
                    cmd.Dispose();
                    throw ex;
                }
            }
        }

        /// <summary>
        /// 执行查询sql语句，返回一个无参数dataset对象
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <param name="paras"></param>
        /// <returns>返回dataset 对象</returns>
        public static DataSet GetDataSetNotPara(string sql)
        {
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(connstr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    //根据传来的参数。决定是sql语句还是存储过程
                    cmd.CommandType = CommandType.StoredProcedure;
                    conn.Open();
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(ds);
                    }
                }
            }
            return ds;
        }

        /// <summary>
        /// 执行查询sql语句，返回一个无参数dataset对象
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <param name="paras"></param>
        /// <returns>返回dataset 对象</returns>
        public static DataTable GetDataTableNotPara(string sql)
        {
            DataTable dt = new DataTable();
            using (SqlConnection conn = new SqlConnection(connstr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    //根据传来的参数。决定是sql语句还是存储过程

                    conn.Open();
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(dt);
                    }
                }
            }
            return dt;
        }
        public static DataTable GetDataTable(string sql, params SqlParameter[] paras)
        {
            DataSet ds = GetDataSet(sql, paras);
            return ds.Tables[0];
        }
        /// <summary>
        /// 执行查询sql语句，返回一个dataset对象
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <param name="paras">查询参数</param>
        /// <returns>返回dataset 对象</returns>
        public static DataSet GetDataSet(string sql, params SqlParameter[] paras)
        {
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(connstr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    //根据传来的参数。决定是sql语句还是存储过程
                    cmd.CommandType = CommandType.Text;
                    //添加参数
                    if (paras != null)
                    { 
                        cmd.Parameters.AddRange(paras);
                    }
                    conn.Open();
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(ds);
                    }
                    conn.Close();
                }
            }
            return ds;
        }
        /// <summary>
        /// 可以执行sql语句或存储过程
        /// </summary>
        /// <param name="text"></param>
        /// <param name="ct"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public static DataTable ProcGetTable(string sql, params SqlParameter[] param)
        {  
            DataSet ds = ProcGetDataSet(sql, param); 
            return ds.Tables[0];
        }
        /// <summary>
        /// 可以执行sql语句或存储过程
        /// </summary>
        /// <param name="text"></param>
        /// <param name="ct"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public static DataSet ProcGetDataSet(string sql, params SqlParameter[] param)
        {
            DataSet ds = new DataSet();

            using (SqlConnection conn = new SqlConnection(connstr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    //根据传来的参数。决定是sql语句还是存储过程 
                    cmd.CommandType = CommandType.StoredProcedure;
                    //添加参数
                    if (param != null)
                    { 
                        cmd.Parameters.AddRange(param);
                        //cmd.Parameters.Add("@name", SqlDbType.NVarChar, 20).Value = param[0];
                        //cmd.Parameters.Add("@pwd", SqlDbType.NVarChar, 20).Value = param[1];
                    }
                    conn.Open();
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(ds);
                    }
                    conn.Close();
                }
            }
            return ds;
        }
        /// <summary>
        /// 实现分页功能
        /// </summary>
        /// <param name="sql">sql语句</param>
        /// <param name="paras">参数数组（显示index页和每页显示条数size）</param>
        /// <returns>查询结果</returns>
        public static DataTable GetParaTable(string sql, params SqlParameter[] paras)
        {
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(connstr))
            {
                using (SqlDataAdapter da = new SqlDataAdapter(sql, conn))
                {
                    if (paras != null)
                    {
                        da.SelectCommand.Parameters.AddRange(paras);
                    }
                    da.SelectCommand.CommandType = CommandType.StoredProcedure;
                    da.Fill(ds);
                }
            }
            return ds.Tables[0];
        }
    }
}
