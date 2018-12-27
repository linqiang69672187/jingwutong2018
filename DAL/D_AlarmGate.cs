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
    /// 数据访问类:AlarmGate
    /// </summary>
    public partial class D_AlarmGate
    {
        public D_AlarmGate()
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
            strSql.Append("select count(1) from AlarmGate");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)
			};
            parameters[0].Value = ID;

            return sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text, parameters)>0;
        }


        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_AlarmGate model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into AlarmGate(");
            strSql.Append("DevType,AlarmType,CommonAlarmGate,UrgencyAlarmGate)");
            strSql.Append(" values (");
            strSql.Append("@DevType,@AlarmType,@CommonAlarmGate,@UrgencyAlarmGate)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
					new SqlParameter("@DevType", SqlDbType.Int,4),
					new SqlParameter("@AlarmType", SqlDbType.Int,4),
					new SqlParameter("@CommonAlarmGate", SqlDbType.Int,4),
					new SqlParameter("@UrgencyAlarmGate", SqlDbType.Int,4)};
            parameters[0].Value = model.DevType;
            parameters[1].Value = model.AlarmType;
            parameters[2].Value = model.CommonAlarmGate;
            parameters[3].Value = model.UrgencyAlarmGate;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            return rows;
        }
        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(int CommonAlarmGate, int UrgencyAlarmGate,string TypeName,string Name,int ID)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update AlarmGate set ");

            strSql.Append("CommonAlarmGate=@CommonAlarmGate,");
            strSql.Append("UrgencyAlarmGate=@UrgencyAlarmGate");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {

					new SqlParameter("@CommonAlarmGate", SqlDbType.Int,4),
					new SqlParameter("@UrgencyAlarmGate", SqlDbType.Int,4),
					new SqlParameter("@ID", SqlDbType.Int,4)};

            parameters[0].Value = Transform.CheckIsNull(CommonAlarmGate);
            parameters[1].Value =Transform.CheckIsNull(UrgencyAlarmGate);
            parameters[2].Value = Transform.CheckIsNull(ID);

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from AlarmGate ");
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
        public Model.M_AlarmGate GetModel(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,DevType,AlarmType,CommonAlarmGate,UrgencyAlarmGate from AlarmGate ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)
			};
            parameters[0].Value = ID;
           Model.M_AlarmGate model = new Model.M_AlarmGate();
           using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
           {
               while (read.Read())
               {
                   if (read["ID"] != null && read["ID"].ToString() != "")
                   {
                       model.ID = int.Parse(read["ID"].ToString());
                   }
                   if (read["DevType"] != null && read["DevType"].ToString() != "")
                   {
                       model.DevType = int.Parse(read["DevType"].ToString());
                   }
                   if (read["AlarmType"] != null && read["AlarmType"].ToString() != "")
                   {
                       model.AlarmType = int.Parse(read["AlarmType"].ToString());
                   }
                   if (read["CommonAlarmGate"] != null && read["CommonAlarmGate"].ToString() != "")
                   {
                       model.CommonAlarmGate = int.Parse(read["CommonAlarmGate"].ToString());
                   }
                   if (read["UrgencyAlarmGate"] != null && read["UrgencyAlarmGate"].ToString() != "")
                   {
                       model.UrgencyAlarmGate = int.Parse(read["UrgencyAlarmGate"].ToString());
                   }

               }


           }
           return model;
        }




        /// <summary>
        /// 获取告警总数
        /// </summary>
        /// <returns></returns>
        public int getcount(int DeviceNmae)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select count(*) from AlarmGate as A  ");

            strSql.Append(" left join  DeviceType  as B on A.DevType=B.ID ");

            strSql.Append(" left join  AlarmType as c on A.AlarmType=C.Id ");

            strSql.Append("  where 1=1  ");



            if (DeviceNmae != -2)
            {
                strSql.Append("  and  A.DevType= " + DeviceNmae);

            }




            return int.Parse(sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text).ToString());
        }









        /// <summary>
        /// 分页排序告警
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(int DeviceNmae, int startRowIndex, int maximumRows)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select A.ID, B.TypeName,C.TypeName as Name,A.CommonAlarmGate,A.UrgencyAlarmGate  from AlarmGate as A ");

            strSql.Append(" left join  DeviceType  as B on A.DevType=B.ID ");

            strSql.Append(" left join  AlarmType as c on A.AlarmType=C.Id ");

            strSql.Append("  where 1=1  ");


            if (DeviceNmae != -2)
            {
                strSql.Append("  and  A.DevType= " + DeviceNmae);

            }



            strSql.Append(" order by DevType asc ");


            SqlParameter[] parameters = { };
            return sqlhelper.ExecuteRead(CommandType.Text, strSql.ToString(), startRowIndex, maximumRows, "device", parameters);

        }



     
      
        #endregion  BasicMethod
  
    }
}
