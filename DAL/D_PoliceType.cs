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
    /// 数据访问类:PoliceType
    /// </summary>
    public partial class D_PoliceType
    {
        public D_PoliceType()
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
            strSql.Append("select count(1) from PoliceType");
            strSql.Append(" where ID=@ID ");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)};
				
            parameters[0].Value = ID;
            return sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text, parameters) > 0;
        }


        /// <summary>
        /// 根据警员名称查找警员相对应类型
        /// </summary>
        /// <param name="JYLXName"></param>
        /// <returns></returns>
        public object ExistsJYLX(string JYLXName)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select ID from PoliceType");
            strSql.Append(" where TypeName=@TypeName ");
            SqlParameter[] parameters = {
					new SqlParameter("@TypeName", SqlDbType.VarChar,50)};

            parameters[0].Value = JYLXName;
            return sqlhelper.ExecuteScalarOne(strSql.ToString(), CommandType.Text, parameters);


        }




        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_PoliceType model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into PoliceType(");
            strSql.Append("ID,TypeName)");
            strSql.Append(" values (");
            strSql.Append("@ID,@TypeName)");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4),
					new SqlParameter("@TypeName", SqlDbType.VarChar,50)};
            parameters[0].Value = model.ID;
            parameters[1].Value = model.TypeName;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            return rows;
        }
        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_PoliceType model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update PoliceType set ");
            strSql.Append("ID=@ID,");
            strSql.Append("TypeName=@TypeName");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4),
					new SqlParameter("@TypeName", SqlDbType.VarChar,50)};
            parameters[0].Value = model.ID;
            parameters[1].Value = model.TypeName;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from PoliceType ");
            strSql.Append(" where ID=@ID ");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4),
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
        public Model.M_PoliceType GetModel(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,TypeName from PoliceType ");
            strSql.Append(" where ID=@ID ");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)};
            parameters[0].Value = ID;
    

          Model.M_PoliceType model = new Model.M_PoliceType();

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

              }
          }

          return model;

        }



        /// <summary>
        /// 获得所有列表
        /// </summary>
        /// <returns></returns>
        public DataTable GetList()
        {
            string sql = "select * from PoliceType ";
            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }

        #endregion  BasicMethod
      
    }
}
