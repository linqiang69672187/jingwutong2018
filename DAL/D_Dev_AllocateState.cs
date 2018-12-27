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
    /// 数据访问类:Dev_AllocateState
    /// </summary>
    public partial class D_Dev_AllocateState
    {
        public D_Dev_AllocateState()
        { }

        //实例化
        SqlHelper sqlhelper = new SqlHelper();
        #region  BasicMethod

        /// <summary>
        /// 加载下拉列表
        /// </summary>
        public DataTable GetList()
        {
        
           string sql= "select * from Dev_AllocateState";

            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }


        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_Dev_AllocateState model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into Dev_AllocateState(");
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
        public int Update(Model.M_Dev_AllocateState model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Dev_AllocateState set ");
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
            strSql.Append("delete from Dev_AllocateState ");
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


        #endregion  BasicMethod
   
    }
}
