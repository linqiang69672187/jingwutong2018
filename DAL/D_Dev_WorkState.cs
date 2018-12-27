using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{

    /// <summary>
    /// 数据访问类:Dev_WorkState
    /// </summary>
    public partial class D_Dev_WorkState
    {
        public D_Dev_WorkState()
        { }

        //实例化
        SqlHelper sqlhelper = new SqlHelper();
        #region  BasicMethod

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_Dev_WorkState model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into Dev_WorkState(");
            strSql.Append("ID,StateName)");
            strSql.Append(" values (");
            strSql.Append("@ID,@StateName)");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4),
					new SqlParameter("@StateName", SqlDbType.VarChar,50)};
            parameters[0].Value = model.ID;
            parameters[1].Value = model.StateName;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            return rows;
        }
        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_Dev_WorkState model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Dev_WorkState set ");
            strSql.Append("ID=@ID,");
            strSql.Append("StateName=@StateName");
            strSql.Append(" where ID=@ID and StateName=@StateName ");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4),
					new SqlParameter("@StateName", SqlDbType.VarChar,50)};
            parameters[0].Value = model.ID;
            parameters[1].Value = model.StateName;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(int ID, string StateName)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from Dev_WorkState ");
            strSql.Append(" where ID=@ID and StateName=@StateName ");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4),
					new SqlParameter("@StateName", SqlDbType.VarChar,50)			};
            parameters[0].Value = ID;
            parameters[1].Value = StateName;

      

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            if (rows > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        /// <summary>
        /// 获得所有列表
        /// </summary>
        /// <returns></returns>
        public DataTable GetList()
        {
            string sql = "select * from Dev_WorkState";
            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }
        /// <summary>
        /// 获得所有列表  只要上下线
        /// </summary>
        /// <returns></returns>
        public DataTable GetListLog()
        {
            string sql = "select * from Dev_WorkState where ID in (11,12) ";
            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }



        /// <summary>
        /// 获得所有列表 维修统计
        /// </summary>
        /// <returns></returns>
        public DataTable GetListrEpairs()
        {
            string sql = "select * from Dev_WorkState where ID in (10,20,30,40)";
            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }



        /// <summary>
        /// 查找对应ID
        /// </summary>
        /// <returns></returns>
        public object GetListID(string StateName)
        {
            string sql = "  select ID from Dev_WorkState where StateName= '" + StateName + "'";

            return sqlhelper.ExecuteScalarOne(sql, CommandType.Text);


        }


        #endregion  BasicMethod

    }
}
