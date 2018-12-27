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
    /// 数据访问类:ProjectDetail
    /// </summary>
    public partial class D_ProjectDetail
    {
        public D_ProjectDetail()
        { }
        //实例化
        SqlHelper sqlhelper = new SqlHelper();
        #region  BasicMethod

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_ProjectDetail model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into ProjectDetail(");
            strSql.Append("XMBH,SBXH,SBGG,Price)");
            strSql.Append(" values (");
            strSql.Append("@XMBH,@SBXH,@SBGG,@Price)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
					new SqlParameter("@XMBH", SqlDbType.VarChar,50),
					new SqlParameter("@SBXH", SqlDbType.VarChar,50),
					new SqlParameter("@SBGG", SqlDbType.VarChar,50),
					new SqlParameter("@Price", SqlDbType.Int,4)};
            parameters[0].Value = model.XMBH;
            parameters[1].Value = model.SBXH;
            parameters[2].Value = model.SBGG;
            parameters[3].Value = model.Price;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }
        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int  Update(Model.M_ProjectDetail model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update ProjectDetail set ");
            strSql.Append("XMBH=@XMBH,");
            strSql.Append("SBXH=@SBXH,");
            strSql.Append("SBGG=@SBGG,");
            strSql.Append("Price=@Price");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@XMBH", SqlDbType.VarChar,50),
					new SqlParameter("@SBXH", SqlDbType.VarChar,50),
					new SqlParameter("@SBGG", SqlDbType.VarChar,50),
					new SqlParameter("@Price", SqlDbType.Int,4),
					new SqlParameter("@ID", SqlDbType.Int,4)};
            parameters[0].Value = model.XMBH;
            parameters[1].Value = model.SBXH;
            parameters[2].Value = model.SBGG;
            parameters[3].Value = model.Price;
            parameters[4].Value = model.ID;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from ProjectDetail ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)
			};
            parameters[0].Value = ID;
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
        /// 得到一个对象实体
        /// </summary>
        public Model.M_ProjectDetail GetModel(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,XMBH,SBXH,SBGG,Price from ProjectDetail ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)
			};
            parameters[0].Value = ID;

           Model.M_ProjectDetail model = new Model.M_ProjectDetail();
           using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
           {
               while (read.Read())
               {

                   if (read["ID"] != null && read["ID"].ToString() != "")
                   {
                       model.ID = int.Parse(read["ID"].ToString());
                   }
                   if (read["XMBH"] != null)
                   {
                       model.XMBH = read["XMBH"].ToString();
                   }
                   if (read["SBXH"] != null)
                   {
                       model.SBXH = read["SBXH"].ToString();
                   }
                   if (read["SBGG"] != null)
                   {
                       model.SBGG = read["SBGG"].ToString();
                   }
                   if (read["Price"] != null && read["Price"].ToString() != "")
                   {
                       model.Price = int.Parse(read["Price"].ToString());
                   }


               }
           }
           return model;
        }




        /// <summary>
        /// 获得数据列表
        /// </summary>
        public DataTable GetList(string XMBH)
        {
            string sql = "select * from ProjectDetail where XMBH='" + XMBH + "'";
            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }


        #endregion  BasicMethod

    }
}
