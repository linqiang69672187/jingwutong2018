using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TOOL;

namespace DAL
{
    /// <summary>
    /// 数据访问类:OperationLog
    /// </summary>
    public partial class D_OperationLog
    {
        public D_OperationLog()
        { }

        //实例化
        SqlHelper sqlhelper = new SqlHelper();
        #region  BasicMethod

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_OperationLog model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into OperationLog(");
            strSql.Append("JYBH,Module,OperContent,IpAddr,LogTime,BZ,OptObject)");
            strSql.Append(" values (");
            strSql.Append("@JYBH,@Module,@OperContent,@IpAddr,@LogTime,@BZ,@OptObject)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
					new SqlParameter("@JYBH", SqlDbType.VarChar,50),
					new SqlParameter("@Module", SqlDbType.VarChar,2),
					new SqlParameter("@OperContent", SqlDbType.VarChar,2),
					new SqlParameter("@IpAddr", SqlDbType.VarChar,50),
					new SqlParameter("@LogTime", SqlDbType.DateTime),
                    new SqlParameter("@BZ", SqlDbType.Text),
					new SqlParameter("@OptObject", SqlDbType.VarChar,100)};

            parameters[0].Value =Transform.CheckIsNull(model.JYBH);
            parameters[1].Value =Transform.CheckIsNull(model.Module);
            parameters[2].Value =Transform.CheckIsNull(model.OperContent);
            parameters[3].Value =Transform.CheckIsNull(model.IpAddr);
            parameters[4].Value =Transform.CheckIsNull (model.LogTime);
            parameters[5].Value =Transform.CheckIsNull ( model.BZ);
            parameters[6].Value = Transform.CheckIsNull (model.OptObject);

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            return rows;
        }
        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_OperationLog model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update OperationLog set ");
            strSql.Append("JYBH=@JYBH,");
            strSql.Append("Module=@Module,");
            strSql.Append("OperContent=@OperContent,");
            strSql.Append("IpAddr=@IpAddr,");
            strSql.Append("LogTime=@LogTime");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@JYBH", SqlDbType.VarChar,50),
					new SqlParameter("@Module", SqlDbType.VarChar,2),
					new SqlParameter("@OperContent", SqlDbType.VarChar,2),
					new SqlParameter("@IpAddr", SqlDbType.VarChar,50),
					new SqlParameter("@LogTime", SqlDbType.DateTime),
					new SqlParameter("@ID", SqlDbType.BigInt,8)};
            parameters[0].Value = model.JYBH;
            parameters[1].Value = model.Module;
            parameters[2].Value = model.OperContent;
            parameters[3].Value = model.IpAddr;
            parameters[4].Value = model.LogTime;
            parameters[5].Value = model.ID;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(long ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from OperationLog ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.BigInt)
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
        public  Model.M_OperationLog GetModel(long ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,JYBH,Module,OperContent,IpAddr,LogTime from OperationLog ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.BigInt)
			};
            parameters[0].Value = ID;

            Model.M_OperationLog model = new Model.M_OperationLog();
            using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
            {
                while (read.Read())
                {
                    if (read["ID"] != null && read["ID"].ToString() != "")
                    {
                        model.ID = long.Parse(read["ID"].ToString());
                    }
                    if (read["JYBH"] != null)
                    {
                        model.JYBH = read["JYBH"].ToString();
                    }
                    if (read["Module"] != null)
                    {
                        model.Module = read["Module"].ToString();
                    }
                    if (read["OperContent"] != null)
                    {
                        model.OperContent = read["OperContent"].ToString();
                    }
                    if (read["IpAddr"] != null)
                    {
                        model.IpAddr = read["IpAddr"].ToString();
                    }
                    if (read["LogTime"] != null && read["LogTime"].ToString() != "")
                    {
                        model.LogTime = DateTime.Parse(read["LogTime"].ToString());
                    }


                }

            }

            return model;
           
        }





        /// <summary>
        /// 获取设备总数量 系统日志
        /// </summary>
        /// <returns></returns>
        public int getcount(string three, string second,string search)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select count(A.ID) from OperationLog as A   ");

            strSql.Append(" left join Log_OperateType as B on A.OperContent=B.OperateId  ");
            strSql.Append("  left join Log_Module as C on  A.Module=C.MoudleId  ");
            strSql.Append("  left join ACL_USER as D on A.JYBH=D.JYBH    ");

            strSql.Append("  left join Entity as E  on D.BMDM=E.BMDM  where 1=1  ");
     



            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and D.JYBH like '%" + search + "%'");

            }

            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and D.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (D.BMDM='" + second + "'");

                strSql.Append("  or E.SJBM='" + second + "')");
            }



            return int.Parse(sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text).ToString());
        }





        /// <summary>
        /// 分页排序设备信息 系统日志
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(string three, string second,string search, int startRowIndex, int maximumRows)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select A.ID, D.XM,B.OperateType,C.MoudleName,A.IpAddr,A.LogTime,A.OptObject,A.BZ from OperationLog as A  ");

            strSql.Append(" left join Log_OperateType as B on A.OperContent=B.OperateId  ");
            strSql.Append("  left join Log_Module as C on  A.Module=C.MoudleId  ");
            strSql.Append("  left join ACL_USER as D on A.JYBH=D.JYBH   ");

            strSql.Append("  left join Entity as E  on D.BMDM=E.BMDM  where 1=1  ");



            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and A.JYBH like '%" + search + "%'");

            }

            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and D.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and( D.BMDM='" + second + "'");

                strSql.Append("  or E.SJBM='" + second + "')");
            }







            strSql.Append("  order by LogTime desc");

            SqlParameter[] parameters = { };
            return sqlhelper.ExecuteRead(CommandType.Text, strSql.ToString(), startRowIndex, maximumRows, "device", parameters);

        }



     

        #endregion  BasicMethod

    }
}
