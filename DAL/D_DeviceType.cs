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
    /// 数据访问类:DeviceType
    /// </summary>
    public partial class D_DeviceType
    {
        public D_DeviceType()
        { }

        //实例化
        SqlHelper sqlhelper = new SqlHelper();
        #region  BasicMethod
        /// <summary>
        /// 是否存在该记录
        /// </summary>
        public bool Exists(int ID)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select count(1) from DeviceType");
            strSql.Append(" where ID=@ID ");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)			};
            parameters[0].Value = ID;

            return sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text, parameters) > 0;
        }


        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_DeviceType model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into DeviceType(");
            strSql.Append("ID,TypeName,TypePoto)");
            strSql.Append(" values (");
            strSql.Append("@ID,@TypeName,@TypePoto)");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4),
					new SqlParameter("@TypeName", SqlDbType.VarChar,20),
					new SqlParameter("@TypePoto", SqlDbType.VarChar,100)};
            parameters[0].Value = model.ID;
            parameters[1].Value = model.TypeName;
            parameters[2].Value = model.TypePoto;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }
        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_DeviceType model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update DeviceType set ");
            strSql.Append("TypeName=@TypeName,");
            strSql.Append("TypePoto=@TypePoto");
            strSql.Append(" where ID=@ID ");
            SqlParameter[] parameters = {
					new SqlParameter("@TypeName", SqlDbType.VarChar,20),
					new SqlParameter("@TypePoto", SqlDbType.VarChar,100),
					new SqlParameter("@ID", SqlDbType.Int,4)};
            parameters[0].Value = model.TypeName;
            parameters[1].Value = model.TypePoto;
            parameters[2].Value = model.ID;


            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from DeviceType ");
            strSql.Append(" where ID=@ID ");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)			};
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
        public Model.M_DeviceType GetModel(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,TypeName,TypePoto from DeviceType ");
            strSql.Append(" where ID=@ID ");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)			};
            parameters[0].Value = ID;

          Model.M_DeviceType model = new Model.M_DeviceType();
          using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
          {
              while (read.Read())
              {

                  if (read["ID"] != null && read["ID"].ToString() != "")
                  {
                      model.ID = int.Parse(read["ID"].ToString());
                  }
                  if (read["TypeName"] != null)
                  {
                      model.TypeName = read["TypeName"].ToString();
                  }
                  if (read["TypePoto"] != null)
                  {
                      model.TypePoto = read["TypePoto"].ToString();
                  }

              }
          }

          return model;
        }

        /// <summary>
        /// 模糊查询
        /// </summary>
        /// <returns></returns>
        public DataTable GetLike(string field, string fieldvalue)
        {

            string strSql = "select *  from DeviceType where " + field + "like'" + fieldvalue + "%'";

            DataTable ds = sqlhelper.ExecuteTable(strSql, CommandType.Text);

            return ds;
        }


        /// <summary>
        /// 获得所有列表
        /// </summary>
        /// <returns></returns>
        public DataTable GetList()
        {
            string sql = "select * from DeviceType order by Sort";
            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }

        /// <summary>
        /// 查找对应ID
        /// </summary>
        /// <returns></returns>
        public object GetListID(string TypeName)
        {
            string sql = "  select ID from DeviceType where TypeName like '%" + TypeName+"%'";

            return sqlhelper.ExecuteScalarOne(sql, CommandType.Text) ;
         

        }




        #endregion  BasicMethod
      
    }
}
