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
    /// 数据访问类:Device
    /// </summary>
    public partial class Device
    {
        public Device()
        { }

        //实例化
        SqlHelper sqlhelper = new SqlHelper();

        #region  BasicMethod
        /// <summary>
        /// 是否存在该记录
        /// </summary>
        public bool Exists(string DevId)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select count(1) from Device");
            strSql.Append(" where DevId=@DevId ");
            SqlParameter[] parameters = {
					new SqlParameter("@DevId", SqlDbType.VarChar,50)			};
            parameters[0].Value = DevId;

            return sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text, parameters) > 0;
        }


        /// <summary>
        /// 是否存在该警员编号记录
        /// </summary>
        public bool ExistsJYBH(string JYBH)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select count(1) from Device");
            strSql.Append(" where JYBH=@JYBH  and  JYBH!='' ");
            SqlParameter[] parameters = {
					     new SqlParameter("@JYBH", SqlDbType.VarChar,50),		};
            parameters[0].Value = JYBH;

            return sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text, parameters) > 0;
        }





        /// <summary>
        /// 是否存在该警员编号记录2
        /// </summary>
        public bool ExistsJYBH2(string JYBH,long ID)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select count(1) from Device");
            strSql.Append(" where JYBH=@JYBH  and  JYBH!='' and ID=@ID  ");
            SqlParameter[] parameters = {
					     new SqlParameter("@JYBH", SqlDbType.VarChar,50),
                         new SqlParameter("@ID", SqlDbType.BigInt,8)   
                                        };
            parameters[0].Value = JYBH;
            parameters[1].Value = ID;

        
            return sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text, parameters) > 0;
        }



        
        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.Device model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into Device(");
            strSql.Append("DevType,Manufacturer,DevId,SBXH,SBGG,ProjName,ProjNum,Price,XMFZR,XMFZRDH,CGSJ,BXQ,BFQX,WorkState,BCJYSJ,XCJYSJ,JYBH,BMDM,CreatDate,AllocateState,Cartype,PlateNumber,IMEI,SIMID)");
            strSql.Append(" values (");
            strSql.Append("@DevType,@Manufacturer,@DevId,@SBXH,@SBGG,@ProjName,@ProjNum,@Price,@XMFZR,@XMFZRDH,@CGSJ,@BXQ,@BFQX,@WorkState,@BCJYSJ,@XCJYSJ,@JYBH,@BMDM,@CreatDate,@AllocateState,@Cartype,@PlateNumber,@IMEI,@SIMID)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
                    new SqlParameter("@DevType", SqlDbType.Int,4),
                    new SqlParameter("@Manufacturer", SqlDbType.VarChar,50),
                    new SqlParameter("@DevId", SqlDbType.VarChar,50),
                    new SqlParameter("@SBXH", SqlDbType.VarChar,50),
                    new SqlParameter("@SBGG", SqlDbType.VarChar,50),
                    new SqlParameter("@ProjName", SqlDbType.VarChar,50),
		            new SqlParameter("@ProjNum", SqlDbType.VarChar,50),
                    new SqlParameter("@Price", SqlDbType.VarChar,50),
                    new SqlParameter("@XMFZR", SqlDbType.VarChar,50),
					new SqlParameter("@XMFZRDH", SqlDbType.VarChar,50),
					new SqlParameter("@CGSJ", SqlDbType.Date,3),
					new SqlParameter("@BXQ", SqlDbType. Date,3),
					new SqlParameter("@BFQX", SqlDbType.Date,3),
                    new SqlParameter("@WorkState", SqlDbType.Int,4),
                    new SqlParameter("@BCJYSJ", SqlDbType.DateTime),
                    new SqlParameter("@XCJYSJ", SqlDbType.DateTime),
                    new SqlParameter("@JYBH", SqlDbType.VarChar,50),
                    new SqlParameter("@BMDM", SqlDbType.VarChar,50),
                   new SqlParameter("@CreatDate", SqlDbType.DateTime),
                 new SqlParameter("@AllocateState", SqlDbType.Int,4),

                    new SqlParameter("@Cartype", SqlDbType.VarChar,50),
                    new SqlParameter("@PlateNumber", SqlDbType.VarChar,50),
                
                    //new SqlParameter("@AllocateTime", SqlDbType.DateTime),
                  
                    new SqlParameter("@IMEI", SqlDbType.VarChar,50),
                    new SqlParameter("@SIMID", SqlDbType.VarChar,50),
                   
               
                     };

            parameters[0].Value = model.DevType;
            parameters[1].Value = model.Manufacturer;
            parameters[2].Value = model.DevId;
            parameters[3].Value = model.SBXH;
            parameters[4].Value = model.SBGG;
            parameters[5].Value = model.ProjName;
            parameters[6].Value = model.ProjNum;
            parameters[7].Value = model.Price;
            parameters[8].Value = model.XMFZR;
            parameters[9].Value = model.XMFZRDH;
            parameters[10].Value =Transform.CheckIsNull(model.CGSJ);
            parameters[11].Value =Transform.CheckIsNull(model.BXQ);
            parameters[12].Value =Transform.CheckIsNull(model.BFQX);
            parameters[13].Value = model.WorkState;
            parameters[14].Value =Transform.CheckIsNull(model.BCJYSJ);
            parameters[15].Value = Transform.CheckIsNull(model.XCJYSJ);
            parameters[16].Value = model.JYBH;
            parameters[17].Value = Transform.CheckIsNull(model.BMDM);
            parameters[18].Value = model.CreatDate;

            parameters[19].Value = model.AllocateState;
            parameters[20].Value = Transform.CheckIsNull(model.Cartype);
            parameters[21].Value =Transform.CheckIsNull(model.PlateNumber);
            parameters[22].Value = Transform.CheckIsNull(model.IMEI);
            parameters[23].Value = Transform.CheckIsNull(model.SIMID);


            //parameters[16].Value = model.Price;
            //parameters[17].Value = model.XMFZR;
            //parameters[18].Value = model.XMFZRDH;
            //parameters[19].Value = model.CGSJ;
            //parameters[20].Value = model.BXQ;
            //parameters[21].Value = model.BFQX;

           

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            return rows;
        }

         /// <summary>
        /// 当导入重复时改为更新数据
        /// </summary>
        public int UpdateData(Model.Device model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Device set ");
            strSql.Append("DevType=@DevType,");
            strSql.Append("Manufacturer=@Manufacturer,");
            //strSql.Append("DevId=@DevId,");
            strSql.Append("SBXH=@SBXH,");
            strSql.Append("SBGG=@SBGG,");
            strSql.Append("ProjName=@ProjName,");
            strSql.Append("ProjNum=@ProjNum,");
            strSql.Append("Price=@Price,");
            strSql.Append("XMFZR=@XMFZR,");
            strSql.Append("XMFZRDH=@XMFZRDH,");
            strSql.Append("CGSJ=@CGSJ,");
            strSql.Append("BXQ=@BXQ,");
            strSql.Append("BFQX=@BFQX,");
            strSql.Append("WorkState=@WorkState,");
            strSql.Append("BCJYSJ=@BCJYSJ,");
            strSql.Append("XCJYSJ=@XCJYSJ,");
            strSql.Append("JYBH=@JYBH,");
            strSql.Append("BMDM=@BMDM,");
            strSql.Append("CreatDate=@CreatDate,");
            strSql.Append("AllocateState=@AllocateState,");

            strSql.Append("Cartype=@Cartype,");
            strSql.Append("PlateNumber=@PlateNumber,");
            strSql.Append("IMEI=@IMEI,");
            strSql.Append("SIMID=@SIMID ");


            strSql.Append(" where DevId=@DevId");
            SqlParameter[] parameters = {
				   new SqlParameter("@DevType", SqlDbType.Int,4),
                    new SqlParameter("@Manufacturer", SqlDbType.VarChar,50),
                    new SqlParameter("@SBXH", SqlDbType.VarChar,50),
                    new SqlParameter("@SBGG", SqlDbType.VarChar,50),
                    new SqlParameter("@ProjName", SqlDbType.VarChar,50),
		            new SqlParameter("@ProjNum", SqlDbType.VarChar,50),
                    new SqlParameter("@Price", SqlDbType.VarChar,50),
                    new SqlParameter("@XMFZR", SqlDbType.VarChar,50),
					new SqlParameter("@XMFZRDH", SqlDbType.VarChar,50),
					new SqlParameter("@CGSJ", SqlDbType.Date,3),
					new SqlParameter("@BXQ", SqlDbType. Date,3),
					new SqlParameter("@BFQX", SqlDbType.Date,3),
                    new SqlParameter("@WorkState", SqlDbType.Int,4),
                    new SqlParameter("@BCJYSJ", SqlDbType.DateTime),
                    new SqlParameter("@XCJYSJ", SqlDbType.DateTime),
                    new SqlParameter("@JYBH", SqlDbType.VarChar,50),
                    new SqlParameter("@BMDM", SqlDbType.VarChar,50),
                   new SqlParameter("@CreatDate", SqlDbType.DateTime),
                   new SqlParameter("@AllocateState", SqlDbType.Int,4),

                   
                    new SqlParameter("@Cartype", SqlDbType.VarChar,50),
                    new SqlParameter("@PlateNumber", SqlDbType.VarChar,50),
                    new SqlParameter("@IMEI", SqlDbType.VarChar,50),
                    new SqlParameter("@SIMID", SqlDbType.VarChar,50),

                    new SqlParameter("@DevId", SqlDbType.VarChar,50),
                    //new SqlParameter("@ID", SqlDbType.BigInt,8),
                                        };
            parameters[0].Value = model.DevType;
            parameters[1].Value = model.Manufacturer;
            //parameters[2].Value = model.DevId;
            parameters[2].Value = model.SBXH;
            parameters[3].Value = model.SBGG;
            parameters[4].Value = model.ProjName;
            parameters[5].Value = model.ProjNum;
            parameters[6].Value = model.Price;
            parameters[7].Value = model.XMFZR;
            parameters[8].Value = model.XMFZRDH;
            parameters[9].Value =Transform.CheckIsNull( model.CGSJ);
            parameters[10].Value =Transform.CheckIsNull( model.BXQ);
            parameters[11].Value =Transform.CheckIsNull (model.BFQX);
            parameters[12].Value = model.WorkState;
            parameters[13].Value = Transform.CheckIsNull(model.BCJYSJ);
            parameters[14].Value = Transform.CheckIsNull(model.XCJYSJ);
            parameters[15].Value = model.JYBH;
            parameters[16].Value = Transform.CheckIsNull(model.BMDM);
            parameters[17].Value = model.CreatDate;
            parameters[18].Value = model.AllocateState;

            parameters[19].Value = Transform.CheckIsNull(model.Cartype);
            parameters[20].Value = Transform.CheckIsNull(model.PlateNumber);
            parameters[21].Value = Transform.CheckIsNull(model.IMEI);
            parameters[22].Value = Transform.CheckIsNull(model.SIMID);
            //parameters[24].Value = model.ID;
            parameters[23].Value = model.DevId;


            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }


        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.Device model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Device set ");
            strSql.Append("DevType=@DevType,");
            strSql.Append("Manufacturer=@Manufacturer,");
            //strSql.Append("DevId=@DevId,");
            strSql.Append("SBXH=@SBXH,");
            strSql.Append("SBGG=@SBGG,");
            strSql.Append("ProjName=@ProjName,");
            strSql.Append("ProjNum=@ProjNum,");
            strSql.Append("Price=@Price,");
            strSql.Append("XMFZR=@XMFZR,");
            strSql.Append("XMFZRDH=@XMFZRDH,");
            strSql.Append("CGSJ=@CGSJ,");
            strSql.Append("BXQ=@BXQ,");
            strSql.Append("BFQX=@BFQX,");
            strSql.Append("WorkState=@WorkState,");
            strSql.Append("BCJYSJ=@BCJYSJ,");
            strSql.Append("XCJYSJ=@XCJYSJ,");
            strSql.Append("JYBH=@JYBH,");
            strSql.Append("BMDM=@BMDM,");
            strSql.Append("CreatDate=@CreatDate,");
            strSql.Append("AllocateState=@AllocateState");


            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
				   new SqlParameter("@DevType", SqlDbType.Int,4),
                    new SqlParameter("@Manufacturer", SqlDbType.VarChar,50),
                    //new SqlParameter("@DevId", SqlDbType.VarChar,50),
                    new SqlParameter("@SBXH", SqlDbType.VarChar,50),
                    new SqlParameter("@SBGG", SqlDbType.VarChar,50),
                    new SqlParameter("@ProjName", SqlDbType.VarChar,50),
		            new SqlParameter("@ProjNum", SqlDbType.VarChar,50),
                    new SqlParameter("@Price", SqlDbType.VarChar,50),
                    new SqlParameter("@XMFZR", SqlDbType.VarChar,50),
					new SqlParameter("@XMFZRDH", SqlDbType.VarChar,50),
					new SqlParameter("@CGSJ", SqlDbType.Date,3),
					new SqlParameter("@BXQ", SqlDbType. Date,3),
					new SqlParameter("@BFQX", SqlDbType.Date,3),
                    new SqlParameter("@WorkState", SqlDbType.Int,4),
                    new SqlParameter("@BCJYSJ", SqlDbType.DateTime),
                    new SqlParameter("@XCJYSJ", SqlDbType.DateTime),
                    new SqlParameter("@JYBH", SqlDbType.VarChar,50),
                    new SqlParameter("@BMDM", SqlDbType.VarChar,50),
                   new SqlParameter("@CreatDate", SqlDbType.DateTime),
                   new SqlParameter("@AllocateState", SqlDbType.Int,4),
                   	new SqlParameter("@ID", SqlDbType.BigInt,8),
                                        };
            parameters[0].Value = model.DevType;
            parameters[1].Value = model.Manufacturer;
            //parameters[2].Value = model.DevId;
            parameters[2].Value = model.SBXH;
            parameters[3].Value = model.SBGG;
            parameters[4].Value = model.ProjName;
            parameters[5].Value = model.ProjNum;
            parameters[6].Value = model.Price;
            parameters[7].Value = model.XMFZR;
            parameters[8].Value = model.XMFZRDH;
            parameters[9].Value = model.CGSJ;
            parameters[10].Value = model.BXQ;
            parameters[11].Value = model.BFQX;
            parameters[12].Value = model.WorkState;
            parameters[13].Value = Transform.CheckIsNull(model.BCJYSJ);
            parameters[14].Value = Transform.CheckIsNull(model.XCJYSJ);
            parameters[15].Value = model.JYBH;
            parameters[16].Value =Transform.CheckIsNull(model.BMDM);
            parameters[17].Value = model.CreatDate;
            parameters[18].Value = model.AllocateState;
            parameters[19].Value = model.ID;



            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }



        public int UpdateRemind1(int DevType, string Manufacturer, string SBXH, string ProjName, DateTime CGSJ, DateTime BFQX, string Price, DateTime CreatDate, long ID, string TypeName, string Remind)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Device set ");
            strSql.Append("DevType=@DevType,");
            strSql.Append("Manufacturer=@Manufacturer,");
            strSql.Append("SBXH=@SBXH,");
            strSql.Append("ProjName=@ProjName,");
            strSql.Append("CGSJ=@CGSJ,");
            strSql.Append("BFQX=@BFQX,");
            strSql.Append("Price=@Price,");
            strSql.Append("CreatDate=@CreatDate");



            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
				    new SqlParameter("@DevType", SqlDbType.Int,4),
                    new SqlParameter("@Manufacturer", SqlDbType.VarChar,50),
                    new SqlParameter("@SBXH", SqlDbType.VarChar,50),
                    new SqlParameter("@ProjName", SqlDbType.VarChar,50),
		         	new SqlParameter("@CGSJ", SqlDbType.Date,3),
                    new SqlParameter("@BFQX", SqlDbType.Date,3),
                    new SqlParameter("@Price", SqlDbType.VarChar,50),
                    new SqlParameter("@CreatDate", SqlDbType.DateTime),
                   	new SqlParameter("@ID", SqlDbType.BigInt,8)
                                        };
            parameters[0].Value = DevType;
            parameters[1].Value = Transform.CheckIsNull(Manufacturer).ToString();
            parameters[2].Value = SBXH;
            parameters[3].Value = ProjName;
            parameters[4].Value = CGSJ;
            parameters[5].Value = BFQX;
            parameters[6].Value = Price;
            parameters[7].Value =CreatDate;
            parameters[8].Value = ID;
            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }


        /// <summary>
        /// 更新一条数据 设备提醒表
        /// </summary>
        public int UpdateRemind(Model.Device model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Device set ");
            strSql.Append("DevType=@DevType,");
            strSql.Append("Manufacturer=@Manufacturer,");
            strSql.Append("SBXH=@SBXH,");
            strSql.Append("ProjName=@ProjName,");
            strSql.Append("CGSJ=@CGSJ,");
            strSql.Append("BFQX=@BFQX,");
            strSql.Append("Price=@Price,");
            strSql.Append("CreatDate=@CreatDate");



            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
				    new SqlParameter("@DevType", SqlDbType.Int,4),
                    new SqlParameter("@Manufacturer", SqlDbType.VarChar,50),
                    new SqlParameter("@SBXH", SqlDbType.VarChar,50),
                    new SqlParameter("@ProjName", SqlDbType.VarChar,50),
		         	new SqlParameter("@CGSJ", SqlDbType.Date,3),
                    new SqlParameter("@BFQX", SqlDbType.Date,3),
                    new SqlParameter("@Price", SqlDbType.VarChar,50),
                    new SqlParameter("@CreatDate", SqlDbType.DateTime),
                   	new SqlParameter("@ID", SqlDbType.BigInt,8)
                                        };
            parameters[0].Value = model.DevType;
            parameters[1].Value = model.Manufacturer;
            parameters[2].Value = model.SBXH;
            parameters[3].Value = model.ProjName;
            parameters[4].Value = model.CGSJ;
            parameters[5].Value = model.BFQX;
            parameters[6].Value = model.Price;
            parameters[7].Value = model.CreatDate;
            parameters[8].Value = model.ID;
            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }




        /// <summary>
        /// 更新WorkState设备状态 where 值为DevId设备编号
        /// </summary>
        public int UpdateWorkState_DevId(Model.Device model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Device set ");
            ;
            strSql.Append("WorkState=@WorkState ");

            strSql.Append(" where DevId=@DevId");
            SqlParameter[] parameters = {
				
                    new SqlParameter("@WorkState", SqlDbType.Int,4),
        
                   new SqlParameter("@DevId", SqlDbType.VarChar,50),
                                        };
            parameters[0].Value = model.WorkState;
            parameters[1].Value = model.DevId; ;



            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }





        /// <summary>
        /// 更新WorkState设备状态
        /// </summary>
        public int UpdateWorkState(Model.Device model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Device set ");
     ;
            strSql.Append("WorkState=@WorkState ");

            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
				
                    new SqlParameter("@WorkState", SqlDbType.Int,4),
        
                   	new SqlParameter("@ID", SqlDbType.BigInt,8),
                                        };
            parameters[0].Value = model.WorkState;
            parameters[1].Value = model.ID;



            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }







        /// <summary>
        /// 更新AllocateState设备状态
        /// </summary>
        public int UpdateAllocateState(Model.Device model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Device set ");
            strSql.Append("JYBH='',");
            strSql.Append("AllocateState=@AllocateState ");

            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
				
                    new SqlParameter("@AllocateState", SqlDbType.Int,4),
        
                   	new SqlParameter("@ID", SqlDbType.BigInt,8),
                                        };
            parameters[0].Value = model.AllocateState;
            parameters[1].Value = model.ID;



            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }




        /// <summary>
        /// 更新AllocateState设备状态 where 条件是DevId
        /// </summary>
        public int UpdateAllocateState_WDevId(string DevId)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Device set ");
            ;
            strSql.Append("AllocateState=1 ");

            strSql.Append(" where DevId=@DevId");
            SqlParameter[] parameters = {
        
                  new SqlParameter("@DevId", SqlDbType.VarChar,50),
                                        };
            parameters[0].Value = DevId;




            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }




        /// <summary>
        /// 更新JYBH
        /// </summary>
        public int Update_JYBH(Model.Device model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Device set ");
        
            strSql.Append("JYBH=@JYBH ");

            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
				  
                    new SqlParameter("@JYBH", SqlDbType.VarChar,50),
                     	new SqlParameter("@ID", SqlDbType.BigInt,8)
                                        };
            parameters[0].Value = model.JYBH;
            parameters[1].Value = model.ID;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(long ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from Device ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.BigInt)
			};
            parameters[0].Value = ID;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            if (rows > 0)
            {
                //TOOL.Login.OptObject = ID.ToString(); 
             //TOOL.Login.BZ = strSql.ToString();


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
        public Model.Device GetModel(long ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,BMDM,DevType,DevId,Cartype,PlateNumber,WorkState,AllocateState,AllocateTime,JYBH,IMEI,SIMID,Manufacturer,SBXH,SBGG,ProjName,ProjNum,Price,XMFZR,XMFZRDH,CGSJ,BXQ,BFQX,BCJYSJ,XCJYSJ,CreatDate from Device ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.BigInt)
			};
            parameters[0].Value = ID;

            Model.Device model = new Model.Device();
            using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
            {
                while (read.Read())
                {
                    if (read["ID"] != null && read["ID"].ToString() != "")
                    {
                        model.ID = long.Parse(read["ID"].ToString());
                    }
                    if (read["BMDM"] != null)
                    {
                        model.BMDM = read["BMDM"].ToString();
                    }
                    if (read["DevType"] != null && read["DevType"].ToString() != "")
                    {
                        model.DevType = int.Parse(read["DevType"].ToString());
                    }
                    if (read["DevId"] != null)
                    {
                        model.DevId = read["DevId"].ToString();
                    }
                    if (read["Cartype"] != null)
                    {
                        model.Cartype = read["Cartype"].ToString();
                    }
                    if (read["PlateNumber"] != null)
                    {
                        model.PlateNumber = read["PlateNumber"].ToString();
                    }
                    if (read["WorkState"] != null && read["WorkState"].ToString() != "")
                    {
                        model.WorkState = int.Parse(read["WorkState"].ToString());
                    }
                    if (read["AllocateState"] != null && read["AllocateState"].ToString() != "")
                    {
                        model.AllocateState = int.Parse(read["AllocateState"].ToString());
                    }
                    if (read["AllocateTime"] != null && read["AllocateTime"].ToString() != "")
                    {
                        model.AllocateTime = DateTime.Parse(read["AllocateTime"].ToString());
                    }
                    if (read["JYBH"] != null)
                    {
                        model.JYBH = read["JYBH"].ToString();
                    }
                    if (read["IMEI"] != null)
                    {
                        model.IMEI = read["IMEI"].ToString();
                    }
                    if (read["SIMID"] != null)
                    {
                        model.SIMID = read["SIMID"].ToString();
                    }
                    if (read["Manufacturer"] != null)
                    {
                        model.Manufacturer = read["Manufacturer"].ToString();
                    }
                    if (read["SBXH"] != null)
                    {
                        model.SBXH = read["SBXH"].ToString();
                    }
                    if (read["SBGG"] != null)
                    {
                        model.SBGG = read["SBGG"].ToString();
                    }
                    if (read["ProjName"] != null)
                    {
                        model.ProjName = read["ProjName"].ToString();
                    }
                    if (read["ProjNum"] != null)
                    {
                        model.ProjNum = read["ProjNum"].ToString();
                    }
                    if (read["Price"] != null)
                    {
                        model.Price = read["Price"].ToString();
                    }
                    if (read["XMFZR"] != null)
                    {
                        model.XMFZR = read["XMFZR"].ToString();
                    }
                    if (read["XMFZRDH"] != null)
                    {
                        model.XMFZRDH = read["XMFZRDH"].ToString();
                    }
                    if (read["CGSJ"] != null && read["CGSJ"].ToString() != "")
                    {
                        model.CGSJ = DateTime.Parse(read["CGSJ"].ToString());
                    }
                    if (read["BXQ"] != null && read["BXQ"].ToString() != "")
                    {
                        model.BXQ = DateTime.Parse(read["BXQ"].ToString());
                    }
                    if (read["BFQX"] != null && read["BFQX"].ToString() != "")
                    {
                        model.BFQX = DateTime.Parse(read["BFQX"].ToString());
                    }
                    if (read["BCJYSJ"] != null && read["BCJYSJ"].ToString() != "")
                    {
                        model.BCJYSJ = DateTime.Parse(read["BCJYSJ"].ToString());
                    }
                    if (read["XCJYSJ"] != null && read["XCJYSJ"].ToString() != "")
                    {
                        model.XCJYSJ = DateTime.Parse(read["XCJYSJ"].ToString());
                    }
                    if (read["CreatDate"] != null && read["CreatDate"].ToString() != "")
                    {
                        model.CreatDate = DateTime.Parse(read["CreatDate"].ToString());
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

            string strSql = "select *  from Device where " + field + "like'" + fieldvalue + "%'";

            DataTable ds = sqlhelper.ExecuteTable(strSql, CommandType.Text);

            return ds;
        }


        /// <summary>
        /// 获得所有列表
        /// </summary>
        /// <returns></returns>
        public DataTable GetList()
        {
            string sql = "select * from Device";
            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }


        /// <summary>
        /// 获取设备总数量
        /// </summary>
        /// <returns></returns>
        public int getcount(int DeviceNmae, string strat, string now, int State, string three,string second, string search,int AllocateState )
        {
            StringBuilder strSql = new StringBuilder();
            if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("select COUNT( B.TypeName) as TypeName from device as A ");
                strSql.Append(" INNER JOIN DeviceType as B ");
                strSql.Append(" on A.DevType=B.ID  ");
                strSql.Append(" INNER JOIN Entity  as D  on  A.BMDM=D.BMDM  ");
                strSql.Append(" INNER JOIN ACL_USER as C on A.JYBH =C.JYBH   ");
                strSql.Append(" INNER JOIN  Position as E on C.LDJB=E.ID  where 1=1");
            }
            else
            {

                strSql.Append("select COUNT( B.TypeName) as TypeName from device as A ");
                strSql.Append(" left JOIN DeviceType as B ");
                strSql.Append(" on A.DevType=B.ID  ");
                strSql.Append(" left JOIN Entity  as D  on  A.BMDM=D.BMDM  ");
                strSql.Append(" left JOIN ACL_USER as C on A.JYBH =C.JYBH  ");
                strSql.Append(" left JOIN Position as E on C.LDJB=E.ID  where 1=1");
            }

            if (DeviceNmae != -2)
            {

                strSql.Append(" and A.DevType=" + DeviceNmae);
            }
            if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
            {
                strSql.Append(" and  A.CGSJ  BETWEEN '" + Convert.ToDateTime(strat) + "' and  '" + Convert.ToDateTime(now) + "'");
            }

            if (State != -2)
            {
                strSql.Append(" and A.WorkState=" + State);
            }

            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and A.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (A.BMDM='" + second + "'");

                strSql.Append("  or D.SJBM='" + second + "')");//上级部门编号
            }





            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and ( A.DevId like '%" + search + "%'  or   C.XM like '%" + search + "%' or C.JYBH like '%" + search + "%') ");
            }


                if (AllocateState != -2)
            {
                strSql.Append(" and A.AllocateState= " + AllocateState);//回收状态
            }




            return int.Parse(sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text).ToString());
        }

            


        /// <summary>
        /// 分页排序设备信息
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(int DeviceNmae, string strat, string now, int State, string three,string second,string search,int AllocateState, int startRowIndex, int maximumRows)
        {
            StringBuilder strSql = new StringBuilder();
            if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("select A.DevId,C.XM,A.DevType,A.ID,B.TypeName,SBXH,DevId,Manufacturer,SBGG,A.CGSJ,ProjNum,XMFZR,XMFZRDH,A.JYBH,A.BMDM,A.WorkState,A.AllocateState  from device as A ");
                strSql.Append(" INNER JOIN DeviceType as B ");
                strSql.Append(" on A.DevType=B.ID  ");
                strSql.Append(" INNER JOIN Entity  as D  on  A.BMDM=D.BMDM   ");

                strSql.Append(" INNER JOIN ACL_USER as C on A.JYBH =C.JYBH ");

                strSql.Append(" INNER JOIN  Position as E on C.LDJB=E.ID  where 1=1");
            }


            else
            {
                strSql.Append("select A.DevId,C.XM,A.DevType,A.ID,B.TypeName,SBXH,DevId,Manufacturer,SBGG,A.CGSJ,ProjNum,XMFZR,XMFZRDH,A.JYBH,A.BMDM,A.WorkState,A.AllocateState  from device as A ");
                strSql.Append(" left JOIN DeviceType as B ");
                strSql.Append(" on A.DevType=B.ID  ");
                strSql.Append(" left JOIN Entity  as D  on  A.BMDM=D.BMDM   ");

                strSql.Append(" left JOIN ACL_USER as C on A.JYBH =C.JYBH ");

                strSql.Append(" left JOIN  Position as E on C.LDJB=E.ID  where 1=1");



            }


            if (DeviceNmae != -2)
            {

                strSql.Append(" and A.DevType=" + DeviceNmae);
            }
            if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
            {
                strSql.Append(" and  A.CGSJ  BETWEEN '" + Convert.ToDateTime(strat) + "' and  '" + Convert.ToDateTime(now) + "'");
            }

            if (State != -2)
            {
                strSql.Append(" and A.WorkState=" + State);
            }

            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and A.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (A.BMDM='" + second + "'");

                strSql.Append("  or D.SJBM='" + second + "')");//上级部门编号
            }




            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and ( A.DevId like '%" + search + "%'  or   C.XM like '%" + search + "%' or C.JYBH like '%" + search + "%') ");
            }

            if (AllocateState != -2)
            {
                strSql.Append(" and A.AllocateState= " + AllocateState);//回收状态
            }

      
                strSql.Append(" ORDER BY D.Sort desc,E.Weight desc,A.JYBH desc,A.DevType desc ");


            SqlParameter[] parameters = { };
            return sqlhelper.ExecuteRead(CommandType.Text, strSql.ToString(), startRowIndex, maximumRows, "device", parameters);

        }



/// <summary>
        /// 导出设备登记数据使用
/// </summary>
/// <param name="DeviceNmae"></param>
/// <param name="strat"></param>
/// <param name="now"></param>
/// <param name="State"></param>
/// <param name="three"></param>
/// <param name="search"></param>
/// <returns></returns>
        public DataTable OutExcel(int DeviceNmae, string strat, string now, int State, string three, string second, string search,int AllocateState)
        {
            StringBuilder strSql = new StringBuilder();
            if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("select * from device as A ");
                strSql.Append(" INNER JOIN DeviceType as B ");
                strSql.Append(" on A.DevType=B.ID ");
                strSql.Append(" INNER JOIN Dev_WorkState  as C on A.WorkState=c.ID  ");
                strSql.Append("  INNER JOIN Entity as D on  A.BMDM=D.BMDM ");
                strSql.Append("   INNER JOIN ACL_USER as E on A.JYBH =E.JYBH  where 1=1 ");
            }
            else
            {
                strSql.Append("select * from device as A ");
                strSql.Append(" left JOIN DeviceType as B ");
                strSql.Append(" on A.DevType=B.ID ");
                strSql.Append(" left JOIN Dev_WorkState  as C on A.WorkState=c.ID  ");
                strSql.Append("  left JOIN Entity as D on  A.BMDM=D.BMDM ");
                strSql.Append("   left JOIN ACL_USER as E on A.JYBH =E.JYBH  where 1=1 ");
            
            }


            if (DeviceNmae != -2)
            {

                strSql.Append(" and A.DevType=" + DeviceNmae);
            }
            if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
            {
                strSql.Append(" and  A.CGSJ  BETWEEN '" + Convert.ToDateTime(strat) + "' and  '" + Convert.ToDateTime(now) + "'");
            }

            if (State != -2)
            {
                strSql.Append(" and A.WorkState=" + State);
            }

            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and A.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and A.BMDM='" + second + "'");

                strSql.Append("  or D.SJBM='" + second + "'");//上级部门编号
            }




            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and ( A.DevId like '%" + search + "%'  or   E.XM like '%" + search + "%' or E.JYBH like '%" + search + "%') ");
            }

                if (AllocateState != -2)
            {
                strSql.Append(" and A.AllocateState= " + AllocateState);//回收状态
            }


                strSql.Append(" ORDER BY D.Sort desc,E.Weight desc,A.JYBH desc,A.DevType desc ");

          
    
            return sqlhelper.ExecuteTable(strSql.ToString(), CommandType.Text);

        }






        /// <summary>
        /// 飘窗
        /// </summary>
        public DataTable BindBM(int DevType, string SJBM)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select ROW_NUMBER() over(order by BMMC) as sum,* from Entity as A ");
            strSql.Append(" left join Device as B on A.BMDM=B.BMDM  ");
            strSql.Append(" left join ACL_USER as D on A.BMDM=D.BMDM    ");
            strSql.Append(" left join ACL_USER as D on A.BMDM=D.BMDM where 1=1   ");

            strSql.Append("  and B.DevType="+DevType+" and (A.SJBM='"+SJBM+"' or A.BMDM='"+SJBM+"')   ");



            return sqlhelper.ExecuteTable(strSql.ToString(), CommandType.Text);
        }






/// <summary>
        /// 导出设备提醒表
/// </summary>
/// <param name="Remind"></param>
/// <param name="DeviceNmae"></param>
/// <param name="strat"></param>
/// <param name="now"></param>
/// <param name="search"></param>
/// <returns></returns>
        public DataTable OutExcelRemind(int Remind, int DeviceNmae, string strat, string now, string three, string second, string search)
        {

            StringBuilder strSql = new StringBuilder();

            if (Remind == 1)
            {

                strSql.Append("select * from (");
                strSql.Append(" select A.CreatDate, A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append(" DATEDIFF(D,GETDATE(),XCJYSJ) as XCJYSJ1, '检定提醒' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append(" ) as newtable");
                strSql.Append(" where newtable.XCJYSJ1<=60");
            }
            else if (Remind == 2)
            {

                strSql.Append("select * from (");
                strSql.Append(" select A.CreatDate, A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append(" DATEDIFF(D,GETDATE(),BFQX) as BFQX1, '报废提醒' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append(" ) as newtable");
                strSql.Append(" where newtable.BFQX1<=180");
            }

            else if (Remind == 3)
            {
                strSql.Append(" select * from (");
                strSql.Append(" select A.CreatDate, A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append(" DATEDIFF(D,GETDATE(),BXQ) as BXQ1, '保修提醒' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append(" ) as newtable");
                strSql.Append(" where newtable.BXQ1<=180");
            }

            else if (Remind == -1)
            {
                strSql.Append(" select * from (");
                strSql.Append(" select A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append(" DATEDIFF(d,GETDATE(),BXQ) as BXQ1, '无' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append(" ) as newtable");
                strSql.Append(" where newtable.WorkState='-1'");

            }
            else if (Remind == 0)
            {
                strSql.Append(" select * from (");
                strSql.Append(" select * from (");
                strSql.Append(" select A.CreatDate, A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append(" DATEDIFF(D,GETDATE(),XCJYSJ) as XCJYSJ1, '检定提醒' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append("  ) as newtable");
                strSql.Append(" where newtable.XCJYSJ1<=60");

                strSql.Append(" UNION");

                strSql.Append(" select * from (");
                strSql.Append(" select A.CreatDate, A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append("  DATEDIFF(D,GETDATE(),BFQX) as BFQX1, '报废提醒' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append("  ) as newtable");
                strSql.Append("  where newtable.BFQX1<=180");

                strSql.Append(" UNION");

                strSql.Append(" select * from (");
                strSql.Append("  select A.CreatDate, A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append("  DATEDIFF(D,GETDATE(),BXQ) as BXQ1, '保修提醒' as Remind from Device as A");
                strSql.Append("  left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append("  ) as newtable");
                strSql.Append("  where newtable.BXQ1<=180");
                strSql.Append(" ) as newtable where 1=1 ");


            }

            if (Remind >0)
            {
                if (DeviceNmae != -2)
                {

                    strSql.Append(" and newtable.DevType=" + DeviceNmae);
                }
                if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
                {
                    strSql.Append(" and  newtable.CGSJ  BETWEEN '" + Convert.ToDateTime(strat) + "' and  '" + Convert.ToDateTime(now) + "'");
                }

                if (!string.IsNullOrEmpty(search))
                {
                    strSql.Append(" and  newtable.DevId like '%" + search + "%'");

                }
                if (!string.IsNullOrEmpty(three))
                {
                    strSql.Append(" and newtable.BMDM= '" + three + "'");
                }

                else if (!string.IsNullOrEmpty(second))
                {
                    strSql.Append("  and (newtable.BMDM='" + second + "'");

                    strSql.Append("  or newtable.SJBM='" + second + "')");//上级部门编号
                }

            }
            else if (Remind != -1)
            {
                if (DeviceNmae != -2)
                {

                    strSql.Append(" and newtable.DevType=" + DeviceNmae);
                }
                if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
                {
                    strSql.Append(" and  newtable.CGSJ  BETWEEN '" + Convert.ToDateTime(strat) + "' and  '" + Convert.ToDateTime(now) + "'");
                }

                if (!string.IsNullOrEmpty(search))
                {
                    strSql.Append(" and  newtable.DevId like '%" + search + "%'");

                }
                if (!string.IsNullOrEmpty(three))
                {
                    strSql.Append(" and newtable.BMDM= '" + three + "'");
                }

                else if (!string.IsNullOrEmpty(second))
                {
                    strSql.Append("  and (newtable.BMDM='" + second + "'");

                    strSql.Append("  or newtable.SJBM='" + second + "')");//上级部门编号
                }

            }


            return sqlhelper.ExecuteTable(strSql.ToString(), CommandType.Text);


        }




        /// <summary>
        /// 获取设备总数量 设备提醒
        /// </summary>
        /// <returns></returns>
        public int getcountRemind(int Remind, int DeviceNmae, string strat, string now, string three, string second, string search)
        {
            StringBuilder strSql = new StringBuilder();

            if (Remind == 1)
            {

                strSql.Append("select COUNT(newtable.WorkState) from (");
                strSql.Append(" select  A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append(" DATEDIFF(D,GETDATE(),XCJYSJ) as XCJYSJ1, '检定提醒' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append(" ) as newtable");
                strSql.Append(" where newtable.XCJYSJ1<=60");
            }

            else if (Remind == 2)
            {

                strSql.Append("select COUNT(newtable.WorkState) from (");
                strSql.Append(" select  A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append(" DATEDIFF(D,GETDATE(),BFQX) as BFQX1, '报废提醒' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append(" ) as newtable");
                strSql.Append(" where newtable.BFQX1<=180");
            }

            else if (Remind == 3)
            {
                strSql.Append("select COUNT(newtable.WorkState) from (");
                strSql.Append(" select  A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append(" DATEDIFF(D,GETDATE(),BXQ) as BXQ1, '保修提醒' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append(" ) as newtable");
                strSql.Append(" where newtable.BXQ1<=180");
            }

            else if (Remind == -1)
            {
                strSql.Append("select COUNT(newtable.WorkState) from (");
                strSql.Append(" select  A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append(" DATEDIFF(d,GETDATE(),BXQ) as BXQ1, '无' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append(" ) as newtable");
                strSql.Append(" where newtable.WorkState='-1'");

            }

            else if (Remind == 0)
            {

                strSql.Append("select(");
                strSql.Append(" (select COUNT(newtable.WorkState)  from (");
                strSql.Append(" select  A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append(" DATEDIFF(D,GETDATE(),XCJYSJ) as XCJYSJ1, '检定提醒' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append("  ) as newtable");
                strSql.Append(" where newtable.XCJYSJ1<=60");

 

                if (DeviceNmae != -2)
                {

                    strSql.Append(" and newtable.DevType=" + DeviceNmae);
                }


                if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
                {
                    strSql.Append(" and  newtable.CGSJ  BETWEEN '" + Convert.ToDateTime(strat) + "' and  '" + Convert.ToDateTime(now) + "'");
                }
                if (!string.IsNullOrEmpty(search))
                {
                    strSql.Append(" and  newtable.DevId like '%" + search + "%'");

                }
                if (!string.IsNullOrEmpty(three))
                {
                    strSql.Append(" and newtable.BMDM= '" + three + "'");
                }

                else if (!string.IsNullOrEmpty(second))
                {
                    strSql.Append("  and (newtable.BMDM='" + second + "'");

                    strSql.Append("  or newtable.SJBM='" + second + "')");//上级部门编号
                }

                strSql.Append(")");

                strSql.Append(" +");

                strSql.Append(" (select COUNT(newtable.WorkState)  from (");
                strSql.Append(" select  A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append(" DATEDIFF(D,GETDATE(),BFQX) as BFQX1, '报废提醒' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append(" ) as newtable");
                strSql.Append(" where newtable.BFQX1<=180");


                if (DeviceNmae != -2)
                {

                    strSql.Append(" and newtable.DevType=" + DeviceNmae);
                }


                if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
                {
                    strSql.Append(" and  newtable.CGSJ  BETWEEN '" + Convert.ToDateTime(strat) + "' and  '" + Convert.ToDateTime(now) + "'");
                }
                if (!string.IsNullOrEmpty(search))
                {
                    strSql.Append(" and  newtable.DevId like '%" + search + "%'");

                }
                if (!string.IsNullOrEmpty(three))
                {
                    strSql.Append(" and newtable.BMDM= '" + three + "'");
                }

                else if (!string.IsNullOrEmpty(second))
                {
                    strSql.Append("  and (newtable.BMDM='" + second + "'");

                    strSql.Append("  or newtable.SJBM='" + second + "')");//上级部门编号
                }

                strSql.Append(" )");

                strSql.Append(" +");

                strSql.Append(" (select COUNT(newtable.WorkState)  from (");
                strSql.Append(" select  A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
                strSql.Append(" DATEDIFF(D,GETDATE(),BXQ) as BXQ1, '保修提醒' as Remind from Device as A");
                strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
                strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
                strSql.Append(" ) as newtable");
                strSql.Append(" where newtable.BXQ1<=180");

                if (DeviceNmae != -2)
                {

                    strSql.Append(" and newtable.DevType=" + DeviceNmae);
                }
                if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
                {
                    strSql.Append(" and  newtable.CGSJ  BETWEEN '" + Convert.ToDateTime(strat) + "' and  '" + Convert.ToDateTime(now) + "'");
                }
                if (!string.IsNullOrEmpty(search))
                {
                    strSql.Append(" and  newtable.DevId like '%" + search + "%'");

                }
                if (!string.IsNullOrEmpty(three))
                {
                    strSql.Append(" and newtable.BMDM= '" + three + "'");
                }

                else if (!string.IsNullOrEmpty(second))
                {
                    strSql.Append("  and (newtable.BMDM='" + second + "'");

                    strSql.Append("  or newtable.SJBM='" + second + "')");//上级部门编号
                }
                strSql.Append(")");

                strSql.Append(")");


            }


            if (Remind >0)
            {
                if (DeviceNmae != -2)
                {

                    strSql.Append(" and newtable.DevType=" + DeviceNmae);
                }
                if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
                {
                    strSql.Append(" and  newtable.CGSJ  BETWEEN '" + Convert.ToDateTime(strat) + "' and  '" + Convert.ToDateTime(now) + "'");
                }

                if (!string.IsNullOrEmpty(search))
                {
                    strSql.Append(" and  newtable.DevId like '%" + search + "%'");

                }

                if (!string.IsNullOrEmpty(three))
                {
                    strSql.Append(" and newtable.BMDM= '" + three + "'");
                }

                else if (!string.IsNullOrEmpty(second))
                {
                    strSql.Append("  and (newtable.BMDM='" + second + "'");

                    strSql.Append("  or newtable.SJBM='" + second + "')");//上级部门编号
                }



            }
       
   



            return int.Parse(sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text).ToString());
        }



        /// <summary>
        /// 分页排序设备信息 设备提醒
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpagingRemind(int Remind, int DeviceNmae, string strat, string now,string three,string second,string search, int startRowIndex, int maximumRows)
        {
            StringBuilder strSql = new StringBuilder();

           if(Remind==1)
            { 

            strSql.Append("select * from (");
            strSql.Append(" select  A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
            strSql.Append(" DATEDIFF(D,GETDATE(),XCJYSJ) as XCJYSJ1, '检定提醒' as Remind from Device as A");
            strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
            strSql.Append(" left join Entity as C on A.BMDM=c.BMDM");
            strSql.Append(" ) as newtable");
            strSql.Append(" where newtable.XCJYSJ1<=60");
            }
            else if(Remind==2)
            { 

            strSql.Append("select * from (");
            strSql.Append(" select A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
            strSql.Append(" DATEDIFF(D,GETDATE(),BFQX) as BFQX1, '报废提醒' as Remind from Device as A");
            strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
            strSql.Append(" left join Entity as C on A.BMDM=c.BMDM");
            strSql.Append(" ) as newtable");
            strSql.Append(" where newtable.BFQX1<=180");
            }

            else if(Remind==3)
            { 
            strSql.Append(" select * from (");
            strSql.Append(" select A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
            strSql.Append(" DATEDIFF(D,GETDATE(),BXQ) as BXQ1, '保修提醒' as Remind from Device as A");
            strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
            strSql.Append(" left join Entity as C on A.BMDM=c.BMDM");
            strSql.Append(" ) as newtable");
            strSql.Append(" where newtable.BXQ1<=180");
            }
           else if (Remind == -1)
           {
               strSql.Append(" select * from (");
               strSql.Append(" select A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
               strSql.Append(" DATEDIFF(d,GETDATE(),BXQ) as BXQ1, '无' as Remind from Device as A");
               strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
               strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
               strSql.Append(" ) as newtable");
               strSql.Append(" where newtable.WorkState='-1'");
           
           }

           else if (Remind == 0)
           {
               strSql.Append(" select * from (");
               strSql.Append(" select * from (");
               strSql.Append(" select A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
               strSql.Append(" DATEDIFF(D,GETDATE(),XCJYSJ) as XCJYSJ1, '检定提醒' as Remind from Device as A");
               strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
               strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
               strSql.Append("  ) as newtable");
               strSql.Append(" where newtable.XCJYSJ1<=60");

               strSql.Append(" UNION");

               strSql.Append(" select * from (");
               strSql.Append(" select A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
               strSql.Append("  DATEDIFF(D,GETDATE(),BFQX) as BFQX1, '报废提醒' as Remind from Device as A");
               strSql.Append(" left join DeviceType as B on A.DevType=B.ID");
               strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
               strSql.Append("  ) as newtable");
               strSql.Append("  where newtable.BFQX1<=180");

               strSql.Append(" UNION");

               strSql.Append(" select * from (");
               strSql.Append("  select A.DevId, A.WorkState, A.DevType,A.ID,B.TypeName,A.Manufacturer,A.SBXH,A.SBGG,A.ProjName,A.CGSJ,A.BFQX,A.Price,C.SJBM,C.BMDM,");
               strSql.Append("  DATEDIFF(D,GETDATE(),BXQ) as BXQ1, '保修提醒' as Remind from Device as A");
               strSql.Append("  left join DeviceType as B on A.DevType=B.ID");
               strSql.Append(" left join Entity as C on A.BMDM=C.BMDM");
               strSql.Append("  ) as newtable");
               strSql.Append("  where newtable.BXQ1<=180");
               strSql.Append(" ) as newtable where 1=1 ");


           }

           if (Remind>0)
           {


               if (DeviceNmae != -2)
               {

                   strSql.Append(" and newtable.DevType=" + DeviceNmae);
               }



            if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
            {
                strSql.Append(" and  newtable.CGSJ  BETWEEN '" + Convert.ToDateTime(strat) + "' and  '" + Convert.ToDateTime(now) + "'");
            }

            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and  newtable.DevId '"+search+"'");
            
            }



            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and newtable.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (newtable.BMDM='" + second + "'");

                strSql.Append("  or newtable.SJBM='" + second + "')");//上级部门编号
            }



           }
            else if(Remind!=-1)
           {
               if (DeviceNmae != -2)
               {

                   strSql.Append(" and newtable.DevType=" + DeviceNmae);
               }
               if (!string.IsNullOrEmpty(strat) && !string.IsNullOrEmpty(now))
               {
                   strSql.Append(" and  newtable.CGSJ  BETWEEN '" + Convert.ToDateTime(strat) + "' and  '" + Convert.ToDateTime(now) + "'");
               }

               if (!string.IsNullOrEmpty(search))
               {
                   strSql.Append(" and  newtable.DevId like '%" + search + "%'");

               }


               if (!string.IsNullOrEmpty(three))
               {
                   strSql.Append(" and newtable.BMDM= '" + three + "'");
               }

               else if (!string.IsNullOrEmpty(second))
               {
                   strSql.Append("  and (newtable.BMDM='" + second + "'");

                   strSql.Append("  or newtable.SJBM='" + second + "')");//上级部门编号
               }

           }

           
   


            SqlParameter[] parameters = { };
            return sqlhelper.ExecuteRead(CommandType.Text, strSql.ToString(), startRowIndex, maximumRows, "device", parameters);

        }





        #endregion  BasicMethod

    }
}
