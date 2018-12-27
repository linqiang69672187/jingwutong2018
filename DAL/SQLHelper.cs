using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;


namespace DAL
{
    /// <summary>
    /// sqlserver数据库操作类
    /// </summary>
    public  class SqlHelper
    {
        private  string connectionString;
        public string ConnectionString
        {
            set { connectionString = value; }
        }

 

        public SqlHelper()
        {
            connectionString = ConfigurationManager.ConnectionStrings["ConnectionStringName"].ConnectionString;//@"Server=(local); DataBase=jyb_db;uid=sa;pwd=a123456";
        }
        public SqlHelper(string connstring)
        {
            connectionString = connstring;
        }

        public DataTable ExecuteDataTable(string sql)
        {
            return ExecuteTable(sql, CommandType.Text, null);
        }

        public DataTable ExecuteTable(string sql, CommandType commandtype)
        {
            return ExecuteTable(sql, commandtype, null);
        }




        /// <summary>
        /// 执行查询，返回datatable
        /// </summary>
        /// <param name="sql">要执行的sql语句</param>
        /// <param name="commandtype">要执行的查询语句类型，如存储过程或SQL文本命令</param>
        /// <param name="paramters">Transact-SQL语句或存储过程的参数数组</param>
        /// <returns></returns>ExecuteTable
        public DataTable ExecuteTable(string sql, CommandType commandtype, SqlParameter[] parameters)
        {
            DataTable dt = new DataTable();
            using (SqlConnection sqlconnection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(sql, sqlconnection))
                {
                    command.CommandType = commandtype;
                    if (parameters != null)
                    {
                        foreach (SqlParameter parameter in parameters)
                        {
                            command.Parameters.Add(parameter);
                        }
                    }
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    adapter.Fill(dt);
                }
            }
            return dt;
        }




        #region 分页
        public  DataTable ExecuteRead(CommandType cmdtype, string cmdText, int startRowIndex, int maximumRows, string tableName, params SqlParameter[] Parameters)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlDataAdapter dr = new SqlDataAdapter(cmdText, conn);
                DataSet ds = new DataSet();
                if (Parameters != null)
                {
                    foreach (SqlParameter Parameter in Parameters)
                    {
                        dr.SelectCommand.Parameters.Add(Parameter);
                    }
                }
                dr.Fill(ds, startRowIndex, maximumRows, tableName);
                return ds.Tables[0];
            }
        }
        #endregion




        /// <summary>
        /// 执行查询返回dataset
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public DataSet Executetable(string sql)
        { 
       SqlConnection sqlconnection = new SqlConnection(connectionString);
        SqlDataAdapter da=new SqlDataAdapter(sql,sqlconnection);
        DataSet ds = new DataSet();
        da.Fill(ds);
        return ds;

        }
        /// <summary>
        /// 执行查询返回一个SqlDataReader
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="commandtype"></param>
        /// <param name="parameter"></param>
        /// <returns></returns>
        public SqlDataReader ExecuteSqlReader(string sql, CommandType commandtype, SqlParameter[] parameters)
        {
            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(sql, conn);
            if (parameters != null)
            {
                foreach (SqlParameter parameter in parameters)
                {
                    cmd.Parameters.Add(parameter);
                }
            }
            conn.Open();
            return cmd.ExecuteReader(CommandBehavior.CloseConnection);
        }

        public SqlDataReader ExecuteSqlReader(string sql, CommandType commandtype)
        {
            return ExecuteSqlReader(sql, commandtype, null);
        }

        public SqlDataReader ExecuteSqlReader(string sql)
        {
            return ExecuteSqlReader(sql, CommandType.Text, null);
        }



        /// <summary>
        /// 判断是否有数据
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="commandType"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public int ExecuteScalar(string sql, CommandType commandType, SqlParameter[] parameters)
        {
            int result; ;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    command.CommandType = commandType;
                    if (parameters != null)
                    {
                        foreach (SqlParameter parameter in parameters)
                        {
                            command.Parameters.Add(parameter);
                        }
                    }
                    connection.Open();//打开数据库连接
                    result =int.Parse(command.ExecuteScalar().ToString());
                }
            }
            return result;//返回查询结果的第一行第一列，忽略其它行和列
        }




        /// <summary>
        /// 执行查询返回第一条数据
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="commandType"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public Object ExecuteScalarOne(string sql, CommandType commandType, SqlParameter[] parameters)
        {
            object result = null;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    command.CommandType = commandType;
                    if (parameters != null)
                    {
                        foreach (SqlParameter parameter in parameters)
                        {
                            command.Parameters.Add(parameter);
                        }
                    }
                    connection.Open();//打开数据库连接
                    result = command.ExecuteScalar();
                }
            }
            return result;//返回查询结果的第一行第一列，忽略其它行和列
        }


        public Object ExecuteScalar(string sql, CommandType commandType)
        {
            return ExecuteScalar(sql, commandType, null);
        }

        public Object ExecuteScalar(string sql)
        {
            return ExecuteScalar(sql, CommandType.Text, null);
        }





        public Object ExecuteScalarOne(string sql, CommandType commandType)
        {
            return ExecuteScalarOne(sql, commandType, null);
        }

        public Object ExecuteScalarOne(string sql)
        {
            return ExecuteScalarOne(sql, CommandType.Text, null);
        }
        /// <summary>
        /// 对数据库执行更新和插入操作
        /// </summary>
        /// <param name="sql">要执行的SQL语句</param>·
        /// <param name="commandType">要执行的查询语句的类型，如存储过程或者SQL文本命令</param>
        /// <param name="parameters">Transact-SQL 语句或存储过程的参数数组</param>
        /// <returns></returns>
        public int ExecuteNonQuery(string sql, CommandType commandType, SqlParameter[] parameters)
        {
            int count = 0;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(sql, connection))
                {
                    command.CommandType = commandType;//设置command的CommandType为指定的CommandType
                    //如果同时传入了参数，则添加这些参数
                    if (parameters != null)
                    {
                        command.Parameters.AddRange(parameters);
                        //foreach (SqlParameter parameter in parameters)
                        //{
                        //    command.Parameters.Add(parameter);
                        //}
                    }
                    connection.Open();//打开数据库连接
                    count = command.ExecuteNonQuery();
                }
            }
            return count;//返回执行增删改操作之后，数据库中受影响的行数
        }

        public int ExecuteNonQuery(string sql)
        {
            return ExecuteNonQuery(sql, CommandType.Text, null);
        }

        public int ExecuteNonQuery(string sql, CommandType commandType)
        {
            return ExecuteNonQuery(sql, commandType, null);
        }

        /// <summary>
        /// 返回当前连接的数据库中所有由用户创建的数据库
        /// </summary>
        /// <returns></returns>
        public DataTable GetTables()
        {
            DataTable data = null;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();//打开数据库连接
                data = connection.GetSchema("Tables");
            }
            return data;
        }

        public byte CheckToken(string id, string token) //0 正常 1 token不存在 2 token不匹配，登录状态丢失
        {
            if (string.IsNullOrEmpty(token))
            {
                return 1;   
            }
            string sql = @"SELECT User_Token FROM User_Security WHERE User_ID = " + id;
            object tmp = ExecuteScalar(sql);
            if (tmp == null)
            {
                return 1;
            }

            if (tmp.ToString().Trim() == token)
            {
                return 0;
            }
            else
            {
                return 2;
            }        
        }

        public bool SetToken(string sql)
        {
            int i = ExecuteNonQuery(sql);
            if (i > 0)
            {
                return true;
            }
            return false;
        }


    }
}