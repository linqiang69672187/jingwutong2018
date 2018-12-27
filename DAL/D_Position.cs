using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{/// <summary>
    /// 数据访问类:Position
    /// </summary>
    public partial class D_Position
    {
        public D_Position()
        { }
        //实例化
        SqlHelper sqlhelper = new SqlHelper();
        #region  BasicMethod


        /// <summary>
        /// 获得所有列表
        /// </summary>
        /// <returns></returns>
        public DataTable GetList()
        {
            string sql = "select * from Position order by Weight desc";
            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }

        /// <summary>
        /// 根据职务名称查找职务ID
        /// </summary>
        /// <returns></returns>
        public  object GetListLDJB(string S_PositionName)
        {
            string sql = "select ID from Position where PositionName='"+S_PositionName+"'";
            return sqlhelper.ExecuteScalarOne(sql, CommandType.Text);

        }


        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_Position model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into Position(");
            strSql.Append("PositionName,Description,Weight)");
            strSql.Append(" values (");
            strSql.Append("@PositionName,@Description,@Weight)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
					new SqlParameter("@PositionName", SqlDbType.VarChar,50),
					new SqlParameter("@Description", SqlDbType.VarChar,100),
					new SqlParameter("@Weight", SqlDbType.Int,4)};
            parameters[0].Value = model.PositionName;
            parameters[1].Value = model.Description;
            parameters[2].Value = model.Weight;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            return rows;
        }
        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_Position model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Position set ");
            strSql.Append("PositionName=@PositionName,");
            strSql.Append("Description=@Description,");
            strSql.Append("Weight=@Weight");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@PositionName", SqlDbType.VarChar,50),
					new SqlParameter("@Description", SqlDbType.VarChar,100),
					new SqlParameter("@Weight", SqlDbType.Int,4),
					new SqlParameter("@ID", SqlDbType.Int,4)};
            parameters[0].Value = model.PositionName;
            parameters[1].Value = model.Description;
            parameters[2].Value = model.Weight;
            parameters[3].Value = model.ID;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from Position ");
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
        public Model.M_Position GetModel(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,PositionName,Description,Weight from Position ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)
			};
            parameters[0].Value = ID;

         Model.M_Position model = new Model.M_Position();

         using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
         {
             while (read.Read())
             {

                 if (read["ID"] != null && read["ID"].ToString() != "")
                 {
                     model.ID = int.Parse(read["ID"].ToString());
                 }
                 if (read["PositionName"] != null)
                 {
                     model.PositionName = read["PositionName"].ToString();
                 }
                 if (read["Description"] != null)
                 {
                     model.Description = read["Description"].ToString();
                 }
                 if (read["Weight"] != null && read["Weight"].ToString() != "")
                 {
                     model.Weight = int.Parse(read["Weight"].ToString());
                 }

             }

         }
         return model;
      
        }


        #endregion  BasicMethod
    
    }
}
