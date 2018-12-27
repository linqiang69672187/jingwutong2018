using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace JingWuTong
{
    public class SQLHelper
    {
        private static String connectionString = Program.sqladb;

        #region ExecuteNonQuery封装

        public static int ExecuteNonQuery(CommandType cmdtype, string cmdText, params SqlParameter[] Parameters)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = cmdText;
                    cmd.CommandType = cmdtype;
                    cmd.Connection = conn;
                    foreach (SqlParameter Parameter in Parameters)
                    {
                        cmd.Parameters.Add(Parameter);
                    }
                    return cmd.ExecuteNonQuery();
                }
            }

        }
        #endregion

        #region ExecuteScalar封装
        public static object ExecuteScalar(CommandType cmdtype, string cmdText, params SqlParameter[] Parameters)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = cmdText;
                    cmd.CommandType = cmdtype;
                    cmd.Connection = conn;

                    foreach (SqlParameter Parameter in Parameters)
                    {
                        cmd.Parameters.Add(Parameter);
                    }
                    return cmd.ExecuteScalar();
                }
            }

        }
        #endregion

        #region ExecuteScalar封装
        public static object ExecuteScalarStrProc(CommandType cmdtype, string StoredProcedureName, params SqlParameter[] Parameters)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandType = cmdtype;
                    cmd.CommandText = StoredProcedureName;

                    foreach (SqlParameter Parameter in Parameters)
                    {
                        cmd.Parameters.Add(Parameter);
                    }
                    return cmd.ExecuteScalar();
                }
            }
        }
        #endregion

        #region ExecuteRead封装
        public static DataTable ExecuteRead(CommandType cmdtype, string cmdText, int startRowIndex, int maximumRows, string tableName, params SqlParameter[] Parameters)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlDataAdapter dr = new SqlDataAdapter(cmdText, conn);
                DataSet ds = new DataSet();
                foreach (SqlParameter Parameter in Parameters)
                {
                    dr.SelectCommand.Parameters.Add(Parameter);
                }
                dr.Fill(ds, startRowIndex, maximumRows, tableName);
                return ds.Tables[0];
            }
        }
        #endregion

        #region ExecuteReadStrProc封装
        public static DataTable ExecuteReadStrProc(CommandType cmdtype, string StoredProcedureName, int startRowIndex, int maximumRows, string tableName, params SqlParameter[] p)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlDataAdapter da = new SqlDataAdapter();
                DataSet ds = new DataSet();
                if (p.Length > 0)
                {
                    da.SelectCommand = new SqlCommand();
                    da.SelectCommand.Connection = conn;
                    da.SelectCommand.CommandText = StoredProcedureName;
                    da.SelectCommand.CommandType = cmdtype;
                    for (int i = 0; i < p.Length; i++)
                    {
                        da.SelectCommand.Parameters.Add(p[i]);
                    }
                    da.Fill(ds, startRowIndex, maximumRows, tableName);
                }
                return ds.Tables[0];
            }
        }
        #endregion
        #region ExecuteReadStrProc封装，不分页
        public static DataTable ExecuteReadStrProc(CommandType cmdtype, string StoredProcedureName, string tableName, params SqlParameter[] p)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlDataAdapter da = new SqlDataAdapter();
                DataSet ds = new DataSet();
                if (p.Length > 0)
                {
                    da.SelectCommand = new SqlCommand();
                    da.SelectCommand.Connection = conn;
                    da.SelectCommand.CommandText = StoredProcedureName;
                    da.SelectCommand.CommandType = cmdtype;
                    for (int i = 0; i < p.Length; i++)
                    {
                        da.SelectCommand.Parameters.Add(p[i]);
                    }
                    da.Fill(ds, tableName);
                }
                return ds.Tables[0];
            }
        }
        #endregion
        #region ExecuteRead封装
        public static DataTable ExecuteRead(CommandType cmdtype, string cmdText, string tableName, params SqlParameter[] Parameters)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlDataAdapter dr = new SqlDataAdapter(cmdText, conn);
                DataSet ds = new DataSet();
                foreach (SqlParameter Parameter in Parameters)
                {
                    dr.SelectCommand.Parameters.Add(Parameter);
                }
                dr.Fill(ds, tableName);
                return ds.Tables[0];
            }
        }
        #endregion

        #region 重写DataReader与ExecuteScalar  QJJ
        public static Object ExecuteScalar(String sql)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = sql;
                    cmd.Connection = con;
                    cmd.Connection.Open();
                    return cmd.ExecuteScalar();
                }
            }
        }


        public static int ExecuteNonQuery(string cmdText)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = cmdText;
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = conn;
                    return cmd.ExecuteNonQuery();
                }
            }

        }

        /// <summary>
        /// 获得单条数据记录
        /// </summary>
        /// <param name="objlist"></param>
        /// <param name="sql"></param>
        /// <param name="cmdtype"></param>
        public static void ExecuteDataReader(ref IList<Object> objlist, String sql, CommandType cmdtype)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = sql;
                    cmd.CommandType = cmdtype;
                    cmd.Connection = con;
                    cmd.Connection.Open();
                    using (SqlDataReader sdr = cmd.ExecuteReader())
                    {
                        if (!sdr.HasRows)
                            return;
                        sdr.Read();
                        for (int i = 0; i < sdr.FieldCount; i++)
                        {
                            objlist.Add(sdr[i]);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 获得所有字段或某个字段值
        /// </summary>
        /// <param name="objlist"></param>
        /// <param name="sql"></param>
        public static void ExecuteDataReader(ref IList<Object> objlist, String sql)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.Connection.Open();
                    SqlDataReader sdr = cmd.ExecuteReader();
                    if (!sdr.HasRows)
                        return;
                    while (sdr.Read())
                    {
                        objlist.Add(sdr[0]);
                    }
                }
            }
        }



        /// <summary>
        /// 获取对应表字段
        /// </summary>
        /// <param name="objlist"></param>
        /// <param name="sql"></param>
        /// <param name="cmdtype"></param>
        public static void ExecuteDataReaderTableFields(ref IList<Object> objlist, String sql, CommandType cmdtype)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = sql;
                    cmd.CommandType = cmdtype;
                    cmd.Connection = con;
                    cmd.Connection.Open();
                    using (SqlDataReader sdr = cmd.ExecuteReader())
                    {
                        if (!sdr.HasRows)
                            return;
                        while (sdr.Read())
                        {
                            objlist.Add(sdr[0]);
                        }
                    }
                }
            }
        }


        public static void ExecuteDataReader(ref DataTable dt, IList<Object> collist, String sql)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = sql;
                    cmd.Connection = con;
                    cmd.Connection.Open();
                    using (SqlDataReader sdr = cmd.ExecuteReader())
                    {
                        if (!sdr.HasRows)
                            return;
                        DataRow dr = null;
                        while (sdr.Read())
                        {
                            dr = dt.NewRow();
                            foreach (var c in collist)
                            {
                                dr[c.ToString()] = sdr[c.ToString()];
                            }
                            dt.Rows.Add(dr);
                        }

                    }
                }
            }
        }

        /// <summary>
        /// datareader生成DataTable
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="sql"></param>
        public static void ExecuteDataReader(ref DataTable dt, String sql)
        {
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.Connection.Open();
                    using (SqlDataReader sdr = cmd.ExecuteReader())
                    {
                        if (!sdr.HasRows)
                            return;
                        int fieldcount = sdr.FieldCount;
                        for (int i = 0; i < fieldcount; i++)
                        {
                            dt.Columns.Add(sdr.GetName(i));
                        }
                        while (sdr.Read())
                        {
                            DataRow dr = dt.NewRow();
                            for (int j = 0; j < fieldcount; j++)
                            {
                                dr[j] = sdr[j];
                            }
                            dt.Rows.Add(dr);
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 返回DataReader
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public static SqlDataReader GetReader(String sql)
        {
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandText = sql;
            cmd.Connection = con;
            cmd.Connection.Open();
            return cmd.ExecuteReader(CommandBehavior.CloseConnection);
        }
        #endregion

    }

}
