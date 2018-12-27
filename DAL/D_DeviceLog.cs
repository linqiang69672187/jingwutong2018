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
    /// 数据访问类:DeviceLog
    /// </summary>
    public partial class D_DeviceLog
    {
        public D_DeviceLog()
        { }

        //实例化
        SqlHelper sqlhelper = new SqlHelper();
        #region  BasicMethod

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_DeviceLog model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into DeviceLog(");
            strSql.Append("LogType,DevType,DevId,DevState,JYBH,BXR,UserName,Tel,BZ,Entity,LogTime,AllocateState)");
            strSql.Append(" values (");
            strSql.Append("@DevType,@DevId,@DevState,@JYBH,@BXR,@UserName,@Tel,@BZ,@Entity,@LogTime,@AllocateState)");

            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
                    new SqlParameter("@LogType", SqlDbType.Int,4),
					new SqlParameter("@DevType", SqlDbType.Int,4),
                    new SqlParameter("@DevId", SqlDbType.VarChar,50),
                    new SqlParameter("@DevState", SqlDbType.Int,4),
                    new SqlParameter("@JYBH", SqlDbType.VarChar,50),
                    new SqlParameter("@BXR", SqlDbType.VarChar,50),
                         new SqlParameter("@UserName", SqlDbType.VarChar,50),
                    new SqlParameter("@Tel", SqlDbType.VarChar,50),
                    new SqlParameter("@BZ", SqlDbType.VarChar,200),
            		new SqlParameter("@Entity", SqlDbType.VarChar,50),
            		new SqlParameter("@LogTime", SqlDbType.DateTime),
                         new SqlParameter("@AllocateState", SqlDbType.Int,4),                   
                                        };

            parameters[0].Value = model.LogType;
            parameters[1].Value =model.DevType;
            parameters[2].Value = model.DevId;
            parameters[3].Value =model.DevState;
            parameters[4].Value =model.JYBH;
            parameters[5].Value =model.BXR;
            parameters[6].Value = model.UserName;
            parameters[7].Value = model.Tel;
            parameters[8].Value = model.BZ;
            parameters[9].Value = model.Entity;
            parameters[10].Value = model.LogTime;
            parameters[11].Value = model.AllocateState;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }



        /// <summary>
        /// 增加一条数据 只增加 字段
        /// </summary>
        public int AddThree(Model.M_DeviceLog model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into DeviceLog(");
            strSql.Append("LogType,BXR,BZ,LogTime,DevType,DevState,DevId,AllocateState,Entity,JYBH)");
            strSql.Append(" values (");
            strSql.Append("@LogType,@BXR,@BZ,@LogTime,@DevType,@DevState,@DevId,@AllocateState,@Entity,@JYBH)");

            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
				  new SqlParameter("@LogType", SqlDbType.Int,4),
                    new SqlParameter("@BXR", SqlDbType.VarChar,50),
                    new SqlParameter("@BZ", SqlDbType.VarChar,200),
            		new SqlParameter("@LogTime", SqlDbType.DateTime),
            	   new SqlParameter("@DevType", SqlDbType.Int,4),
                   new SqlParameter("@DevState", SqlDbType.Int,4),
                  new SqlParameter("@DevId", SqlDbType.VarChar,50),
                 new SqlParameter("@AllocateState", SqlDbType.Int,4),
                new SqlParameter("@Entity", SqlDbType.VarChar,50),
                new SqlParameter("@JYBH", SqlDbType.VarChar,50),
            };

            parameters[0].Value = model.LogType;
            parameters[1].Value = model.BXR;
            parameters[2].Value = model.BZ;
            parameters[3].Value = model.LogTime;
            parameters[4].Value = model.DevType;
            parameters[5].Value = model.DevState;
            parameters[6].Value = model.DevId;
            parameters[7].Value = model.AllocateState;
            parameters[8].Value = model.Entity;
            parameters[9].Value = model.JYBH;


            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }





        /// <summary>
        /// 同意更新 只更新 BZ LogTime DevState BXR
        /// </summary>
        public int UpdateFour(Model.M_DeviceLog model)
        {
            //DevType,DevId,DevState,JYBH,BXR,Tel,BZ,Entity
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update DeviceLog set ");

            strSql.Append("DevState=@DevState,");
            strSql.Append("BXR=@BXR,");
            strSql.Append("BZ=@BZ,");
            strSql.Append("LogTime=@LogTime ");

            strSql.Append(" where ID=@ID ");
            SqlParameter[] parameters = {
	
					new SqlParameter("@DevState", SqlDbType.Int,4),
                    new SqlParameter("@BXR", SqlDbType.VarChar,50),
                    new SqlParameter("@BZ", SqlDbType.VarChar,200),
					new SqlParameter("@LogTime", SqlDbType.DateTime),
                   	new SqlParameter("@ID", SqlDbType.BigInt,8)};


            parameters[0].Value = model.DevState;
            parameters[1].Value = model.BXR;
            parameters[2].Value = model.BZ;
            parameters[3].Value = model.LogTime;
            parameters[4].Value = model.ID;
     

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }




        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_DeviceLog model)
        {
            //DevType,DevId,DevState,JYBH,BXR,Tel,BZ,Entity
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update DeviceLog set ");
            strSql.Append("DevType=@DevType,");
            strSql.Append("DevId=@DevId,");
            strSql.Append("DevState=@DevState,");
            strSql.Append("JYBH=@JYBH,");
            strSql.Append("BXR=@BXR,");
            strSql.Append("Tel=@Tel,");
            strSql.Append("BZ=@BZ,");
            strSql.Append("Entity=@Entity,");
            strSql.Append("LogTime=@LogTime ");

            strSql.Append(" where ID=@ID ");
            SqlParameter[] parameters = {
	
					new SqlParameter("@DevType", SqlDbType.Int,4),
					new SqlParameter("@DevId", SqlDbType.VarChar,50),
					new SqlParameter("@DevState", SqlDbType.Int,4),
                    new SqlParameter("@JYBH", SqlDbType.VarChar,50),
                    new SqlParameter("@BXR", SqlDbType.VarChar,50),
                    new SqlParameter("@Tel", SqlDbType.VarChar,50),
                    new SqlParameter("@BZ", SqlDbType.VarChar,200),
					new SqlParameter("@Entity", SqlDbType.VarChar,50),
					new SqlParameter("@LogTime", SqlDbType.DateTime),
                   	new SqlParameter("@ID", SqlDbType.BigInt,8)};


            parameters[0].Value = model.DevType;
            parameters[1].Value = model.DevId;
            parameters[2].Value = model.DevState;
            parameters[3].Value = model.JYBH;
            parameters[4].Value = model.BXR;
            parameters[5].Value = model.Tel;
            parameters[6].Value = model.BZ;
            parameters[7].Value = model.Entity;
            parameters[8].Value = model.LogTime;
            parameters[9].Value = model.ID;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(long ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from DeviceLog ");
            strSql.Append(" where ID=@ID ");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.BigInt,8)
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
        public Model.M_DeviceLog GetModel(long ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,LogType,DevId,DevType,DevState,UserName,Entity,Tel,JYBH,LogTime,BXR,BZ from DeviceLog ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.BigInt,8),
						};
            parameters[0].Value = ID;

           Model.M_DeviceLog model = new Model.M_DeviceLog();
           using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
           {
               while (read.Read())
               {
                   if (read["ID"] != null && read["ID"].ToString() != "")
                   {
                       model.ID = long.Parse(read["ID"].ToString());
                   }
                   if (read["LogType"] != null && read["LogType"].ToString() != "")
                   {
                       model.LogType = int.Parse(read["LogType"].ToString());
                   }
                   if (read["DevId"] != null)
                   {
                       model.DevId = read["DevId"].ToString();
                   }
                   if (read["DevType"] != null && read["DevType"].ToString() != "")
                   {
                       model.DevType = int.Parse(read["DevType"].ToString());
                   }
                   if (read["DevState"] != null && read["DevState"].ToString() != "")
                   {
                       model.DevState = int.Parse(read["DevState"].ToString());
                   }
                   if (read["UserName"] != null)
                   {
                       model.UserName = read["UserName"].ToString();
                   }
                   if (read["Entity"] != null)
                   {
                       model.Entity = read["Entity"].ToString();
                   }
                   if (read["Tel"] != null)
                   {
                       model.Tel = read["Tel"].ToString();
                   }
                   if (read["JYBH"] != null)
                   {
                       model.JYBH = read["JYBH"].ToString();
                   }
                   if (read["LogTime"] != null && read["LogTime"].ToString() != "")
                   {
                       model.LogTime = DateTime.Parse(read["LogTime"].ToString());
                   }
                   if (read["BXR"] != null)
                   {
                       model.BXR = read["BXR"].ToString();
                   }
                   if (read["BZ"] != null)
                   {
                       model.BZ = read["BZ"].ToString();
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
            string sql = "select * from DeviceLog";
            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }





       /// <summary>
       /// 导出维修统计
       /// </summary>
       /// <param name="three"></param>
       /// <param name="DeviceNmae"></param>
       /// <param name="strat"></param>
       /// <param name="now"></param>
       /// <param name="State"></param>
       /// <param name="search"></param>
       /// <returns></returns>

        public DataTable OutExcel(string three, string second, int DeviceNmae, string strat, string now, int State, string search)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select A.LogType, A.ID,A.DevState,C.BMMC,A.BXR,A.Tel,A.JYBH,D.TypeName,B.StateName,A.DevId,A.LogTime,A.BZ from DeviceLog as A  ");

            strSql.Append(" left join Dev_WorkState as B on A.DevState= B.ID ");

            strSql.Append("left join  Entity as C on A.Entity =C.BMDM ");

            strSql.Append("left join DeviceType as D on A.DevType=D.ID  where 1=1 and LogType=1");


            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and C.BMDM='" + three + "'");
            }

            if (!string.IsNullOrEmpty(second))
            {
                strSql.Append(" and C.SJBM='" + second + "'");
            }



            if (DeviceNmae != -2)
            {
                strSql.Append("  and A.DevType=" + DeviceNmae);
            }


            if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
            {
                strSql.Append(" and  A.LogTime  BETWEEN '" + Convert.ToDateTime(strat) + "'  and   '" + Convert.ToDateTime(now) + "'");
            }


            if (State != -2)
            {
                strSql.Append(" and  A.DevState=" + State);
            }

            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and A.DevId like '%" + search + "%'");

            }
            strSql.Append("  and  A.AllocateState=2 ");//配发状态

            strSql.Append(" order by  LogTime desc ");

            return sqlhelper.ExecuteTable(strSql.ToString(), CommandType.Text);


        }




   




        /// <summary>
        /// 获取设备总数量 维修统计
        /// </summary>
        /// <returns></returns>
        public int getcount(string three,string second, int DeviceNmae, string strat, string now, int State, string search)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  count(A.ID) from DeviceLog as A  ");

            strSql.Append(" left join Dev_WorkState as B on A.DevState= B.ID ");

            strSql.Append("left join  Entity as C on A.Entity =C.BMDM ");

            strSql.Append("left join DeviceType as D on A.DevType=D.ID  where 1=1 and LogType=1 ");



            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and C.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (C.BMDM='" + second + "'");

                strSql.Append("  or C.SJBM='" + second + "')");//上级部门编号
            }



            if (DeviceNmae != -2)
            {
                strSql.Append("  and A.DevType=" + DeviceNmae);
            }

            if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
            {
                strSql.Append(" and  A.LogTime  BETWEEN '" + Convert.ToDateTime(strat) + "'  and   '" + Convert.ToDateTime(now) + "'");
            }

            if (State != -2)
            {
                strSql.Append(" and  A.DevState=" + State);
            }

            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and A.DevId like '%" + search + "%'");

            }

            strSql.Append("  and  A.AllocateState=2");//配发状态

            return int.Parse(sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text).ToString());
        }









        /// <summary>
        /// 分页排序设备信息 维修统计
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(string three,string second, int DeviceNmae, string strat, string now, int State,string search, int startRowIndex, int maximumRows)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  A.ID,A.DevState,C.BMMC,C.LXR,A.Tel,A.JYBH,D.TypeName,B.StateName,A.DevId,A.LogTime,A.BXR,A.BZ from DeviceLog as A  ");
          
            strSql.Append(" left join Dev_WorkState as B on A.DevState= B.ID ");

            strSql.Append("left join  Entity as C on A.Entity =C.BMDM ");

            strSql.Append("left join DeviceType as D on A.DevType=D.ID  where 1=1 and LogType=1 ");



            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and C.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (C.BMDM='" + second + "'");

                strSql.Append("  or C.SJBM='" + second + "')");//上级部门编号
            }





            if (DeviceNmae != -2)
            {
                strSql.Append("  and A.DevType=" + DeviceNmae);
            }

            if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
            {
                strSql.Append(" and  A.LogTime  BETWEEN '" + Convert.ToDateTime(strat) + "'  and   '" + Convert.ToDateTime(now) + "'");
            }

            if (State != -2)
            {
                strSql.Append(" and  A.DevState=" + State);
            }

            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and A.DevId like '%" + search + "%'");

            }

            strSql.Append("  and  A.AllocateState=2" );//配发状态

            strSql.Append(" order by  LogTime desc ");

   
            SqlParameter[] parameters = { };
            return sqlhelper.ExecuteRead(CommandType.Text, strSql.ToString(), startRowIndex, maximumRows, "device", parameters);

        }



        /// <summary>
        /// 导出回收统计
        /// </summary>
        /// <param name="three"></param>
        /// <param name="DeviceNmae"></param>
        /// <param name="strat"></param>
        /// <param name="now"></param>
        /// <param name="State"></param>
        /// <param name="search"></param>
        /// <returns></returns>

        public DataTable OutExcelRecycle(string three, string second, string strat, string now, int State, string search)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select A.LogType, A.ID,A.DevState,C.BMJC,C.LXR,C.LXDH,A.JYBH,D.TypeName,A.DevId,A.LogTime,A.BZ from DeviceLog as A  ");

            strSql.Append(" left join Device as B on A.DevId=B.DevId");
            strSql.Append(" left join  Entity as C on A.Entity =C.BMDM ");

            strSql.Append(" left join DeviceType as D on A.DevType=D.ID  ");

            strSql.Append(" left join Dev_WorkState as E on A.DevState=E.ID where 1=1  and LogType=1 ");


            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and C.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (C.BMDM='" + second + "'");

                strSql.Append("  or C.SJBM='" + second + "')");//上级部门编号
            }


            if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
            {
                strSql.Append(" and  A.LogTime  BETWEEN '" + Convert.ToDateTime(strat) + "'  and   '" + Convert.ToDateTime(now) + "'");
            }

            if (State != -2)
            {
                strSql.Append(" and  A.DevState=" + State);
            }

            strSql.Append("  and A.AllocateState=1");//未配发=回收


            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and A.DevId like '%" + search + "%'");

            }

            strSql.Append(" order by A.LogTime DESC");


            return sqlhelper.ExecuteTable(strSql.ToString(), CommandType.Text);

        }



        /// <summary>
        /// 获取设备总数量 回收
        /// </summary>
        /// <returns></returns>
        public int getcountRecycle(string three, string second, string strat, string now, int State, string search)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  count(A.ID) from DeviceLog as A  ");

            strSql.Append(" left join Device as B on A.DevId=B.DevId");

            strSql.Append(" left join  Entity as C on A.Entity =C.BMDM ");

            strSql.Append(" left join DeviceType as D on A.DevType=D.ID  ");
            strSql.Append(" left join Dev_WorkState as E on A.DevState=E.ID  ");
            strSql.Append("  left join ACL_USER as F on A.JYBH=F.JYBH  where 1=1  and LogType=1 ");

            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and C.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (C.BMDM='" + second + "'");

                strSql.Append("  or C.SJBM='" + second + "')");//上级部门编号
            }


            if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
            {
                strSql.Append(" and  A.LogTime  BETWEEN '" + Convert.ToDateTime(strat) + "'  and   '" + Convert.ToDateTime(now) + "'");
            }

            if (State != -2)
            {
                strSql.Append(" and  A.DevState=" + State);
            }

            strSql.Append("  and A.AllocateState=1");//未配发=回收


            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and A.DevId like '%" + search + "%'");

            }

            return int.Parse(sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text).ToString());
        }



  

        /// <summary>
        /// 分页排序设备信息 回收
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpagingRecycle(string three, string second, string strat, string now, int State, string search, int startRowIndex, int maximumRows)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  A.ID,A.DevState,C.BMMC,A.BXR,F.SJ,A.JYBH,D.TypeName,A.DevId,A.LogTime,A.BZ,E.StateName from DeviceLog as A  ");

            strSql.Append(" left join Device as B on A.DevId=B.DevId");
            strSql.Append(" left join  Entity as C on A.Entity =C.BMDM ");

            strSql.Append(" left join DeviceType as D on A.DevType=D.ID ");

            strSql.Append(" left join Dev_WorkState as E on A.DevState=E.ID ");

            strSql.Append("  left join ACL_USER as F on A.JYBH=F.JYBH  where 1=1  and LogType=1 ");

            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and C.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (C.BMDM='" + second + "'");

                strSql.Append("  or C.SJBM='" + second + "')");//上级部门编号
            }


            if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
            {
                strSql.Append(" and  A.LogTime  BETWEEN '" + Convert.ToDateTime(strat) + "'  and   '" + Convert.ToDateTime(now) + "'");
            }

            if (State != -2)
            {
                strSql.Append(" and  A.DevState=" + State);
            }
            strSql.Append("  and A.AllocateState=1");//未配发=回收


            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and A.DevId like '%" + search + "%'");

            }


            strSql.Append(" order by A.LogTime DESC");

            SqlParameter[] parameters = { };
            return sqlhelper.ExecuteRead(CommandType.Text, strSql.ToString(), startRowIndex, maximumRows, "device", parameters);

        }










        /// <summary>
        /// 获取设备总数量 设备日志
        /// </summary>
        /// <returns></returns>
        public int getcountLog(int DeviceNmae, string three, string second, int State, string search)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  count(A.ID) from DeviceLog as A  ");

            strSql.Append(" left join DeviceType as B on A.DevType=B.ID");

            strSql.Append(" left join Dev_WorkState as C on A.DevState=C.ID ");

            strSql.Append("  left join Device as D on A.DevId=D.DevId ");

            strSql.Append("  left join  Entity as E on A.Entity=E.BMDM  where 1=1 and LogType=2");


            if (DeviceNmae != -2)
            {

                strSql.Append(" and A.DevType=" + DeviceNmae);
            }


            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and D.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (D.BMDM='" + second + "'");

                strSql.Append("  or E.SJBM='" + second + "')");//上级部门编号
            }


            if (State != -2)
            {
                strSql.Append(" and  A.DevState=" + State);
            }


            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append("  and  A.DevId  like '%" + search + "%'");

            }

            return int.Parse(sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text).ToString());
        }




        /// <summary>
        /// 分页排序设备信息 设备日志
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpagingLog(int DeviceNmae, string three, string second, int State, string search, int startRowIndex, int maximumRows)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select A.ID,A.DevId,B.TypeName,C.StateName,A.UserName,E.BMJC,A.Tel,A.LogTime from DeviceLog as A  ");

            strSql.Append(" left join DeviceType as B on A.DevType=B.ID");

            strSql.Append(" left join Dev_WorkState as C on A.DevState=C.ID ");

            strSql.Append("  left join Device as D on A.DevId=D.DevId ");

            strSql.Append("  left join  Entity as E on A.Entity=E.BMDM  where 1=1 and LogType=2");


            if (DeviceNmae != -2)
            {

                strSql.Append(" and A.DevType=" + DeviceNmae);
            }


            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and D.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (D.BMDM='" + second + "'");

                strSql.Append("  or E.SJBM='" + second + "')");//上级部门编号
            }

            if (State != -2)
            {
                strSql.Append(" and  A.DevState=" + State);
            }


            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and  A.DevId  like '%" + search + "%'");
            
            }


            SqlParameter[] parameters = { };
            return sqlhelper.ExecuteRead(CommandType.Text, strSql.ToString(), startRowIndex, maximumRows, "device", parameters);

        }



        #endregion  BasicMethod

    }
}
