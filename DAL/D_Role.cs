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
    /// 数据访问类:Role
    /// </summary>
    public partial class D_Role
    {
        public D_Role()
        { }
        //实例化
        SqlHelper sqlhelper = new SqlHelper();
        #region  BasicMethod
        /// <summary>
        /// 是否存在该记录
        /// </summary>
        public object Exists(int ID)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select POWER from Role");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)
			};
            parameters[0].Value = ID;
            return sqlhelper.ExecuteScalarOne(strSql.ToString(), CommandType.Text, parameters);
        }


        /// <summary>
        /// 根据角色名称查找角色ID
        /// </summary>
        /// <param name="JSIDName"></param>
        /// <returns></returns>
        public object ExistsJSID(string JSIDName)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select ID from Role");
            strSql.Append(" where RoleName=@RoleName");
            SqlParameter[] parameters = {
							new SqlParameter("@RoleName", SqlDbType.VarChar,50)
			};
            parameters[0].Value = JSIDName;
            return sqlhelper.ExecuteScalarOne(strSql.ToString(), CommandType.Text, parameters);


        }


        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_Role model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into Role(");
            strSql.Append("RoleName,Power,Status,Bz,Crateator,CreateDate)");
            strSql.Append(" values (");
            strSql.Append("@RoleName,@Power,@Status,@Bz,@Crateator,@CreateDate)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
					new SqlParameter("@RoleName", SqlDbType.VarChar,50),
					new SqlParameter("@Power", SqlDbType.Text),
					new SqlParameter("@Status", SqlDbType.Int,4),
					new SqlParameter("@Bz", SqlDbType.VarChar,150),
					new SqlParameter("@Crateator", SqlDbType.VarChar,50),
					new SqlParameter("@CreateDate", SqlDbType.DateTime)};
            parameters[0].Value = model.RoleName;
            parameters[1].Value = model.Power;
            parameters[2].Value = model.Status;
            parameters[3].Value = model.Bz;
            parameters[4].Value = model.Crateator;
            parameters[5].Value = model.CreateDate;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            return rows;
        }
        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_Role model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Role set ");
            strSql.Append("RoleName=@RoleName,");
            strSql.Append("Power=@Power,");
            strSql.Append("Status=@Status,");
            strSql.Append("Bz=@Bz,");
            strSql.Append("Crateator=@Crateator,");
            strSql.Append("CreateDate=@CreateDate");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@RoleName", SqlDbType.VarChar,50),
					new SqlParameter("@Power", SqlDbType.Text),
					new SqlParameter("@Status", SqlDbType.Int,4),
					new SqlParameter("@Bz", SqlDbType.VarChar,150),
					new SqlParameter("@Crateator", SqlDbType.VarChar,50),
					new SqlParameter("@CreateDate", SqlDbType.DateTime),
					new SqlParameter("@ID", SqlDbType.Int,4)};
            parameters[0].Value = model.RoleName;
            parameters[1].Value = model.Power;
            parameters[2].Value = model.Status;
            parameters[3].Value = model.Bz;
            parameters[4].Value = model.Crateator;
            parameters[5].Value = model.CreateDate;
            parameters[6].Value = model.ID;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from Role ");
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
        public Model.M_Role GetModel(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,RoleName,Power,Status,Bz,Crateator,CreateDate from Role ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)
			};
            parameters[0].Value = ID;

            Model.M_Role model = new Model.M_Role();



            using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
            {

                while (read.Read())
                {

                    if (read["ID"] != null && read["ID"].ToString() != "")
                    {
                        model.ID = int.Parse(read["ID"].ToString());
                    }
                    if (read["RoleName"] != null)
                    {
                        model.RoleName = read["RoleName"].ToString();
                    }
                    if (read["Power"] != null)
                    {
                        model.Power = read["Power"].ToString();
                    }
                    if (read["Status"] != null && read["Status"].ToString() != "")
                    {
                        model.Status = int.Parse(read["Status"].ToString());
                    }
                    if (read["Bz"] != null)
                    {
                        model.Bz = read["Bz"].ToString();
                    }
                    if (read["Crateator"] != null)
                    {
                        model.Crateator = read["Crateator"].ToString();
                    }
                    if (read["CreateDate"] != null && read["CreateDate"].ToString() != "")
                    {
                        model.CreateDate = DateTime.Parse(read["CreateDate"].ToString());
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
            string sql = "select * from Role";
            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }



        /// <summary>
        /// 获取设备总数量
        /// </summary>
        /// <returns></returns>
        public int getcount(string search)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select COUNT(A.ID) from Role as A");

            strSql.Append(" left join ACL_USER as B on A.Crateator=B.JYBH ");

            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" where RoleName like '%" + search + "%'");
            }
            return int.Parse(sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text).ToString());
        }


        /// <summary>
        /// 分页排序设备信息
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(string search, int startRowIndex, int maximumRows)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select A.ID,A.RoleName,B.XM,A.CreateDate, A.Bz from Role as A  ");

            strSql.Append(" left join ACL_USER as B on A.Crateator=B.JYBH ");

            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" where RoleName like '%" + search + "%'");
            }

            SqlParameter[] parameters = { };
            return sqlhelper.ExecuteRead(CommandType.Text, strSql.ToString(), startRowIndex, maximumRows, "device", parameters);

        }

        

        #endregion  BasicMethod
  
    }
}
