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
    /// 数据访问类:Alarm_EveryDayInfo
    /// </summary>
    public partial class D_Alarm_EveryDayInfo
    {
        public D_Alarm_EveryDayInfo()
        { }

        //实例化
        SqlHelper sqlhelper = new SqlHelper();
        #region  BasicMethod



        /// <summary>
        /// 飘窗
        /// </summary>
        public DataTable Bind_Alarm_EveryDayInfo(int DevType, string SJBM,DateTime Start_Time,DateTime End_Time)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("SELECT en.BMDM, en.SJBM as [ParentID],us.XM as [Contacts],de.[DevId],ala.在线时长,ala.[AlarmType] from ");
            strSql.Append(" (");
            strSql.Append("  SELECT [DevId],[AlarmType],sum([Value]) as 在线时长    ");
            strSql.Append("  from [Alarm_EveryDayInfo]     ");
            strSql.Append("  where [AlarmType] <>6 and  [AlarmDay ] >='" + Start_Time + "' and [AlarmDay ] <='" + End_Time + "'   ");
            strSql.Append("  group by [DevId],[AlarmType]   ");
            strSql.Append(" ) as ala ");
            strSql.Append("   left join [Device] as de on de.[DevId] = ala.[DevId] ");
            strSql.Append("    left join [Entity] as en on en.[BMDM] = de.[BMDM] ");
            strSql.Append("    left join ACL_USER as us on de.JYBH = us.JYBH  where 1=1 "  );

            strSql.Append(" and de.[DevType]=" + DevType);

            strSql.Append(" and (en.SJBM='" + SJBM + "' or en.BMDM='" + SJBM + "')   ");

            return sqlhelper.ExecuteTable(strSql.ToString(), CommandType.Text);
        }



        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_Alarm_EveryDayInfo model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into Alarm_EveryDayInfo(");
            strSql.Append("DevId,AlarmDay,AlarmState,AlarmType,Contacts,Entity,Tel,DevType,Value)");
            strSql.Append(" values (");
            strSql.Append("@DevId,@AlarmDay,@AlarmState,@AlarmType,@Contacts,@Entity,@Tel,@DevType,@Value)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
					new SqlParameter("@DevId", SqlDbType.VarChar,50),
					new SqlParameter("@AlarmDay", SqlDbType.Date),
					new SqlParameter("@AlarmState", SqlDbType.Int,4),
					new SqlParameter("@AlarmType", SqlDbType.Int,4),
					new SqlParameter("@Contacts", SqlDbType.VarChar,20),
					new SqlParameter("@Entity", SqlDbType.VarChar,100),
					new SqlParameter("@Tel", SqlDbType.VarChar,50),
					new SqlParameter("@DevType", SqlDbType.Int,4),
					new SqlParameter("@Value", SqlDbType.Int,4)};
            parameters[0].Value = model.DevId;
            parameters[1].Value = model.AlarmDay;
            parameters[2].Value = model.AlarmState;
            parameters[3].Value = model.AlarmType;
            parameters[4].Value = model.Contacts;
            parameters[5].Value = model.Entity;
            parameters[6].Value = model.Tel;
            parameters[7].Value = model.DevType;
            parameters[8].Value = model.Value;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            return rows;
        }
        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_Alarm_EveryDayInfo model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Alarm_EveryDayInfo set ");
            strSql.Append("AlarmDay=@AlarmDay,");
            strSql.Append("AlarmState=@AlarmState,");
            strSql.Append("AlarmType=@AlarmType,");
            strSql.Append("Contacts=@Contacts,");
            strSql.Append("Entity=@Entity,");
            strSql.Append("Tel=@Tel,");
            strSql.Append("DevType=@DevType,");
            strSql.Append("Value=@Value");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@AlarmDay", SqlDbType.Date),
					new SqlParameter("@AlarmState", SqlDbType.Int,4),
					new SqlParameter("@AlarmType", SqlDbType.Int,4),
					new SqlParameter("@Contacts", SqlDbType.VarChar,20),
					new SqlParameter("@Entity", SqlDbType.VarChar,100),
					new SqlParameter("@Tel", SqlDbType.VarChar,50),
					new SqlParameter("@DevType", SqlDbType.Int,4),
					new SqlParameter("@Value", SqlDbType.Int,4),
					new SqlParameter("@ID", SqlDbType.BigInt,8),
					new SqlParameter("@DevId", SqlDbType.VarChar,50)};
            parameters[0].Value = model.AlarmDay;
            parameters[1].Value = model.AlarmState;
            parameters[2].Value = model.AlarmType;
            parameters[3].Value = model.Contacts;
            parameters[4].Value = model.Entity;
            parameters[5].Value = model.Tel;
            parameters[6].Value = model.DevType;
            parameters[7].Value = model.Value;
            parameters[8].Value = model.ID;
            parameters[9].Value = model.DevId;
            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(long ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from Alarm_EveryDayInfo ");
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
        public Model.M_Alarm_EveryDayInfo GetModel(long ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,DevId,AlarmDay,AlarmState,AlarmType,Contacts,Entity,Tel,DevType,Value from Alarm_EveryDayInfo ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.BigInt)
			};
            parameters[0].Value = ID;

            Model.M_Alarm_EveryDayInfo model = new Model.M_Alarm_EveryDayInfo();
              using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
              {
                  while (read.Read())
                  {
                      if (read["ID"] != null && read["ID"].ToString() != "")
                      {
                          model.ID = long.Parse(read["ID"].ToString());
                      }
                      if (read["DevId"] != null)
                      {
                          model.DevId = read["DevId"].ToString();
                      }
                      if (read["AlarmDay"] != null && read["AlarmDay"].ToString() != "")
                      {
                          model.AlarmDay = DateTime.Parse(read["AlarmDay"].ToString());
                      }
                      if (read["AlarmState"] != null && read["AlarmState"].ToString() != "")
                      {
                          model.AlarmState = int.Parse(read["AlarmState"].ToString());
                      }
                      if (read["AlarmType"] != null && read["AlarmType"].ToString() != "")
                      {
                          model.AlarmType = int.Parse(read["AlarmType"].ToString());
                      }
                      if (read["Contacts"] != null)
                      {
                          model.Contacts = read["Contacts"].ToString();
                      }
                      if (read["Entity"] != null)
                      {
                          model.Entity = read["Entity"].ToString();
                      }
                      if (read["Tel"] != null)
                      {
                          model.Tel = read["Tel"].ToString();
                      }
                      if (read["DevType"] != null && read["DevType"].ToString() != "")
                      {
                          model.DevType = int.Parse(read["DevType"].ToString());
                      }
                      if (read["Value"] != null && read["Value"].ToString() != "")
                      {
                          model.Value = int.Parse(read["Value"].ToString());
                      }


                  }
              }
              return model;
           
        }


        

        #endregion  BasicMethod
 
    }
}
