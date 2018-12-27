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
    /// 数据访问类:ACL_USER
    /// </summary>
    public partial class D_ACL_USER
    {
        public D_ACL_USER()
        { }

        //实例化
        SqlHelper sqlhelper = new SqlHelper();

        #region  BasicMethod
        /// <summary>
        /// 用户登录
        /// </summary>
        public object Exists(string JYBH, string password)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select JSID from ACL_USER");
            strSql.Append(" where JYBH=@JYBH and MM=@MM");
            SqlParameter[] parameters = {
               new SqlParameter("@JYBH", SqlDbType.VarChar,50),
				new SqlParameter("@MM", SqlDbType.VarChar,50)

			};
            parameters[0].Value = JYBH;
            parameters[1].Value = password;

            return sqlhelper.ExecuteScalarOne(strSql.ToString(), CommandType.Text, parameters);
        }

        /// <summary>
        /// 用户登录查找姓名
        /// </summary>
        /// <param name="JYBH"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public object ExistsXM(string JYBH, string password)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select XM from ACL_USER");
            strSql.Append(" where JYBH=@JYBH and MM=@MM");
            SqlParameter[] parameters = {
              new SqlParameter("@JYBH", SqlDbType.VarChar,50),
				new SqlParameter("@MM", SqlDbType.VarChar,50)

			};
            parameters[0].Value = JYBH;
            parameters[1].Value = password;

            return sqlhelper.ExecuteScalarOne(strSql.ToString(), CommandType.Text, parameters);
        }


        /// <summary>
        /// 根据警员编号查找姓名
        /// </summary>
        /// <param name="JYBH"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public object ExistsJYXM(string JYBH)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select XM from ACL_USER");
            strSql.Append(" where JYBH=@JYBH ");
            SqlParameter[] parameters = {
              new SqlParameter("@JYBH", SqlDbType.VarChar,50)

			};
            parameters[0].Value = JYBH;
            return sqlhelper.ExecuteScalarOne(strSql.ToString(), CommandType.Text, parameters);
        }





        /// <summary>
        /// 得到一部分对象实体 用于登录操作
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public Model.M_ACL_USER ExistsModel(string JYBH, string password)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("  select A.XM,A.BMDM,B.SJBM from ACL_USER as A ");
            strSql.Append("  left join Entity as B  ");
            strSql.Append("    on A.BMDM=B.BMDM  ");
            strSql.Append(" where A.JYBH=@JYBH and A.MM=@MM");

            SqlParameter[] parameters = {
              new SqlParameter("@JYBH", SqlDbType.VarChar,50),
			  new SqlParameter("@MM", SqlDbType.VarChar,50)
            };
            parameters[0].Value = JYBH;
            parameters[1].Value = password;

            Model.M_ACL_USER model = new Model.M_ACL_USER();
            using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
            {

                while (read.Read())
                {
                    //if (read["ID"] != null && read["ID"].ToString() != "")
                    //{
                    //    model.ID = long.Parse(read["ID"].ToString());
                    //}

                    if (read["XM"] != null)
                    {
                        model.XM = read["XM"].ToString();

                    }
                    if (read["BMDM"] != null)
                    {
                        model.BMDM = read["BMDM"].ToString();
                    }

                    if (read["SJBM"] != null)
                    {
                        model.SJBM = read["SJBM"].ToString();
                    }

                    //if (read["JYBH"] != null)
                    //{
                    //    model.JYBH = read["JYBH"].ToString();
                    //}
                    //if (read["SJ"] != null)
                    //{
                    //    model.SJ = read["SJ"].ToString();
                    //}


                }
            }

            return model;


        }





        /// <summary>
        /// 查找警员编号是否重复
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public bool ExistsJYBH(string JYBH)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select count(1) from ACL_USER");
            strSql.Append(" where JYBH=@JYBH ");
            SqlParameter[] parameters = {
					         new SqlParameter("@JYBH", SqlDbType.VarChar,50)	
                                        
                                        };
            parameters[0].Value = JYBH;

            return sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text, parameters) > 0;
        }




        /// <summary>
        /// 增加一部分数据
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public int AddLittle(Model.M_ACL_USER model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into ACL_USER(");
            strSql.Append("  XM,BMDM,JYBH,SJ  )");
            strSql.Append(" values (");
            strSql.Append("@XM,@BMDM,@JYBH,@SJ)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
                                                    
              		new SqlParameter("@XM", SqlDbType.VarChar,50),                                      
                    new SqlParameter("@BMDM", SqlDbType.VarChar,50),                                
                    new SqlParameter("@JYBH", SqlDbType.VarChar,50),
                    new SqlParameter("@SJ", SqlDbType.VarChar,50)        
                          };

            parameters[0].Value = model.XM;
            parameters[1].Value = model.BMDM;
            parameters[2].Value = model.JYBH;
            parameters[3].Value = model.SJ;
            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            return rows;
        }





        /// <summary>
        /// 警员表新增
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public int AddACL_USER(Model.M_ACL_USER model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into ACL_USER(");
            strSql.Append("  XM,BMDM,JYBH,SJ,SFZMHM,JSID ,JYLX,CJSJ,LDJB)");
            strSql.Append(" values (");
            strSql.Append("@XM,@BMDM,@JYBH,@SJ,@SFZMHM,@JSID ,@JYLX,@CJSJ,@LDJB)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
                                                    
              		new SqlParameter("@XM", SqlDbType.VarChar,50),                                      
                    new SqlParameter("@BMDM", SqlDbType.VarChar,50),                                
                    new SqlParameter("@JYBH", SqlDbType.VarChar,50),
                    new SqlParameter("@SJ", SqlDbType.VarChar,50),    
    				new SqlParameter("@SFZMHM", SqlDbType.VarChar,50),
                    new SqlParameter("@JSID", SqlDbType.Int,4),
                    new SqlParameter("@JYLX", SqlDbType.Int,4),
                    new SqlParameter("@CJSJ", SqlDbType.DateTime),
                    new SqlParameter("@LDJB", SqlDbType.VarChar,2),

                          };

            parameters[0].Value = model.XM;
            parameters[1].Value = model.BMDM;
            parameters[2].Value = model.JYBH;
            parameters[3].Value = model.SJ;
            parameters[4].Value=model.SFZMHM; 
            parameters[5].Value=model.JSID;
            parameters[6].Value =model.JYLX;
            parameters[7].Value=model.CJSJ;
            parameters[8].Value = model.LDJB;
                 
            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            return rows;
        }





        /// <summary>
        /// 增加一条导入数据
        /// </summary>
        public int Add( Model.M_ACL_USER model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into ACL_USER(");
            strSql.Append("JYBH,XM,BMDM,SFZMHM,JTZZ,JYLX,MM,JSID,SJ,BGDH,BZLX,YWGW,SGCLDJ,LDJB,CSRQ,XB,JX,JB,JG,RDTSJ,ZZMM,MZ,XL,ZY,ZW,RDSJ,CGSJ,RXZSJ,ZFZGDJ,ZT,GXSJ,CJSJ)");
            strSql.Append(" values (");
            strSql.Append("@JYBH,@XM,@BMDM,@SFZMHM,@JTZZ,@JYLX,@MM,@JSID,@SJ,@BGDH,@BZLX,@YWGW,@SGCLDJ,@LDJB,@CSRQ,@XB,@JX,@JB,@JG,@RDTSJ,@ZZMM,@MZ,@XL,@ZY,@ZW,@RDSJ,@CGSJ,@RXZSJ,@ZFZGDJ,@ZT,@GXSJ,@CJSJ)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
					new SqlParameter("@JYBH", SqlDbType.VarChar,50),
					new SqlParameter("@XM", SqlDbType.VarChar,50),
					new SqlParameter("@BMDM", SqlDbType.VarChar,50),
					new SqlParameter("@SFZMHM", SqlDbType.VarChar,50),
					new SqlParameter("@JTZZ", SqlDbType.VarChar,150),
					new SqlParameter("@JYLX", SqlDbType.Int,4),
					new SqlParameter("@MM", SqlDbType.VarChar,50),
					new SqlParameter("@JSID", SqlDbType.Int,4),
					new SqlParameter("@SJ", SqlDbType.VarChar,50),
					new SqlParameter("@BGDH", SqlDbType.VarChar,50),
					new SqlParameter("@BZLX", SqlDbType.VarChar,1),
					new SqlParameter("@YWGW", SqlDbType.VarChar,2),
					new SqlParameter("@SGCLDJ", SqlDbType.VarChar,2),
					new SqlParameter("@LDJB", SqlDbType.VarChar,2),
					new SqlParameter("@CSRQ", SqlDbType.DateTime),
					new SqlParameter("@XB", SqlDbType.VarChar,1),
					new SqlParameter("@JX", SqlDbType.VarChar,10),
					new SqlParameter("@JB", SqlDbType.VarChar,50),
					new SqlParameter("@JG", SqlDbType.VarChar,50),
					new SqlParameter("@RDTSJ", SqlDbType.DateTime),
					new SqlParameter("@ZZMM", SqlDbType.VarChar,50),
					new SqlParameter("@MZ", SqlDbType.VarChar,50),
					new SqlParameter("@XL", SqlDbType.VarChar,50),
					new SqlParameter("@ZY", SqlDbType.VarChar,50),
					new SqlParameter("@ZW", SqlDbType.VarChar,50),
					new SqlParameter("@RDSJ", SqlDbType.DateTime),
					new SqlParameter("@CGSJ", SqlDbType.DateTime),
					new SqlParameter("@RXZSJ", SqlDbType.DateTime),
					new SqlParameter("@ZFZGDJ", SqlDbType.VarChar,50),
					new SqlParameter("@ZT", SqlDbType.Int,4),
					new SqlParameter("@GXSJ", SqlDbType.DateTime),
					new SqlParameter("@CJSJ", SqlDbType.DateTime)};
            parameters[0].Value = model.JYBH;
            parameters[1].Value = model.XM;
            parameters[2].Value = model.BMDM;
            parameters[3].Value = model.SFZMHM;
            parameters[4].Value = model.JTZZ;
            parameters[5].Value = model.JYLX;
            parameters[6].Value = model.MM;
            parameters[7].Value = model.JSID;
            parameters[8].Value = model.SJ;
            parameters[9].Value = model.BGDH;
            parameters[10].Value = model.BZLX;
            parameters[11].Value = model.YWGW;
            parameters[12].Value = model.SGCLDJ;
            parameters[13].Value = model.LDJB;
            parameters[14].Value =Transform.CheckIsNull(model.CSRQ);
            parameters[15].Value = model.XB;
            parameters[16].Value = model.JX;
            parameters[17].Value = model.JB;
            parameters[18].Value = model.JG;
            parameters[19].Value =Transform.CheckIsNull( model.RDTSJ);
            parameters[20].Value = model.ZZMM;
            parameters[21].Value = model.MZ;
            parameters[22].Value = model.XL;
            parameters[23].Value = model.ZY;
            parameters[24].Value = model.ZW;
            parameters[25].Value =Transform.CheckIsNull(model.RDSJ);
            parameters[26].Value =Transform.CheckIsNull(model.CGSJ);
            parameters[27].Value =Transform.CheckIsNull(model.RXZSJ);
            parameters[28].Value = model.ZFZGDJ;
            parameters[29].Value = model.ZT;
            parameters[30].Value =Transform.CheckIsNull(model.GXSJ);
            parameters[31].Value =Transform.CheckIsNull(model.CJSJ);

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
                return rows;
         
        }



        /// <summary>
        /// 更新一条导入数据
        /// </summary>
        public int UpdateData(Model.M_ACL_USER model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update ACL_USER set ");
            //strSql.Append("JYBH=@JYBH,");
            strSql.Append("XM=@XM,");
            strSql.Append("BMDM=@BMDM,");
            strSql.Append("SFZMHM=@SFZMHM,");
            strSql.Append("JTZZ=@JTZZ,");
            strSql.Append("JYLX=@JYLX,");
            strSql.Append("MM=@MM,");
            strSql.Append("JSID=@JSID,");
            strSql.Append("SJ=@SJ,");
            strSql.Append("BGDH=@BGDH,");
            strSql.Append("BZLX=@BZLX,");
            strSql.Append("YWGW=@YWGW,");
            strSql.Append("SGCLDJ=@SGCLDJ,");
            strSql.Append("LDJB=@LDJB,");
            strSql.Append("CSRQ=@CSRQ,");
            strSql.Append("XB=@XB,");
            strSql.Append("JX=@JX,");
            strSql.Append("JB=@JB,");
            strSql.Append("JG=@JG,");
            strSql.Append("RDTSJ=@RDTSJ,");
            strSql.Append("ZZMM=@ZZMM,");
            strSql.Append("MZ=@MZ,");
            strSql.Append("XL=@XL,");
            strSql.Append("ZY=@ZY,");
            strSql.Append("ZW=@ZW,");
            strSql.Append("RDSJ=@RDSJ,");
            strSql.Append("CGSJ=@CGSJ,");
            strSql.Append("RXZSJ=@RXZSJ,");
            strSql.Append("ZFZGDJ=@ZFZGDJ,");
            strSql.Append("ZT=@ZT,");
            strSql.Append("GXSJ=@GXSJ,");
            strSql.Append("CJSJ=@CJSJ");
            //strSql.Append(" where ID=@ID");
            strSql.Append(" where JYBH=@JYBH");
            SqlParameter[] parameters = {
                    //new SqlParameter("@JYBH", SqlDbType.VarChar,50),
					new SqlParameter("@XM", SqlDbType.VarChar,50),
					new SqlParameter("@BMDM", SqlDbType.VarChar,50),
					new SqlParameter("@SFZMHM", SqlDbType.VarChar,50),
					new SqlParameter("@JTZZ", SqlDbType.VarChar,150),
					new SqlParameter("@JYLX", SqlDbType.Int,4),
					new SqlParameter("@MM", SqlDbType.VarChar,50),
					new SqlParameter("@JSID", SqlDbType.Int,4),
					new SqlParameter("@SJ", SqlDbType.VarChar,50),
					new SqlParameter("@BGDH", SqlDbType.VarChar,50),
					new SqlParameter("@BZLX", SqlDbType.VarChar,1),
					new SqlParameter("@YWGW", SqlDbType.VarChar,2),
					new SqlParameter("@SGCLDJ", SqlDbType.VarChar,2),
					new SqlParameter("@LDJB", SqlDbType.VarChar,2),
					new SqlParameter("@CSRQ", SqlDbType.DateTime),
					new SqlParameter("@XB", SqlDbType.VarChar,1),
					new SqlParameter("@JX", SqlDbType.VarChar,10),
					new SqlParameter("@JB", SqlDbType.VarChar,50),
					new SqlParameter("@JG", SqlDbType.VarChar,50),
					new SqlParameter("@RDTSJ", SqlDbType.DateTime),
					new SqlParameter("@ZZMM", SqlDbType.VarChar,50),
					new SqlParameter("@MZ", SqlDbType.VarChar,50),
					new SqlParameter("@XL", SqlDbType.VarChar,50),
					new SqlParameter("@ZY", SqlDbType.VarChar,50),
					new SqlParameter("@ZW", SqlDbType.VarChar,50),
					new SqlParameter("@RDSJ", SqlDbType.DateTime),
					new SqlParameter("@CGSJ", SqlDbType.DateTime),
					new SqlParameter("@RXZSJ", SqlDbType.DateTime),
					new SqlParameter("@ZFZGDJ", SqlDbType.VarChar,50),
					new SqlParameter("@ZT", SqlDbType.Int,4),
					new SqlParameter("@GXSJ", SqlDbType.DateTime),
					new SqlParameter("@CJSJ", SqlDbType.DateTime),
                    //new SqlParameter("@ID", SqlDbType.BigInt,8)
                    	new SqlParameter("@JYBH", SqlDbType.VarChar,50)
                                        
                                        };
            //parameters[0].Value = model.JYBH;
            parameters[0].Value = model.XM;
            parameters[1].Value = model.BMDM;
            parameters[2].Value = model.SFZMHM;
            parameters[3].Value = model.JTZZ;
            parameters[4].Value = model.JYLX;
            parameters[5].Value = model.MM;
            parameters[6].Value = model.JSID;
            parameters[7].Value = model.SJ;
            parameters[8].Value = model.BGDH;
            parameters[9].Value = model.BZLX;
            parameters[10].Value = model.YWGW;
            parameters[11].Value = model.SGCLDJ;
            parameters[12].Value = model.LDJB;
            parameters[13].Value = Transform.CheckIsNull(model.CSRQ);
            parameters[14].Value = model.XB;
            parameters[15].Value = model.JX;
            parameters[16].Value = model.JB;
            parameters[17].Value = model.JG;
            parameters[18].Value = Transform.CheckIsNull(model.RDTSJ);
            parameters[19].Value = model.ZZMM;
            parameters[20].Value = model.MZ;
            parameters[21].Value = model.XL;
            parameters[22].Value = model.ZY;
            parameters[23].Value = model.ZW;
            parameters[24].Value = Transform.CheckIsNull(model.RDSJ);
            parameters[25].Value = Transform.CheckIsNull(model.CGSJ);
            parameters[26].Value = Transform.CheckIsNull(model.RXZSJ);
            parameters[27].Value = model.ZFZGDJ;
            parameters[28].Value = model.ZT;
            parameters[29].Value = Transform.CheckIsNull(model.GXSJ);
            parameters[30].Value = Transform.CheckIsNull(model.CJSJ);
            parameters[31].Value = model.JYBH;


            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;

        }





        /// <summary>
        /// 更新一条部分数据
        /// </summary>
        public int UpdatePart(Model.M_ACL_USER model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update ACL_USER set ");
            strSql.Append("XM=@XM,");
            strSql.Append("BMDM=@BMDM,");
            strSql.Append("JYBH=@JYBH,");
            strSql.Append("SJ=@SJ,");
            strSql.Append("GXSJ=@GXSJ");

            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
	                new SqlParameter("@XM", SqlDbType.VarChar,50),
                    new SqlParameter("@BMDM", SqlDbType.VarChar,50),
					new SqlParameter("@JYBH", SqlDbType.VarChar,50),
					new SqlParameter("@SJ", SqlDbType.VarChar,50),
                    new SqlParameter("@GXSJ", SqlDbType.DateTime),
                    new SqlParameter("@ID", SqlDbType.BigInt,8)
				};

            parameters[0].Value = model.XM;
            parameters[1].Value = model.BMDM;
            parameters[2].Value = model.JYBH;
            parameters[3].Value = model.SJ;
            parameters[4].Value = model.GXSJ;
            parameters[5].Value = model.ID;
            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;

        }



        /// <summary>
        ///警员表更新
        /// </summary>
        public int UpdateACL_USE(Model.M_ACL_USER model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update ACL_USER set ");
            strSql.Append("XM=@XM,");
            strSql.Append("BMDM=@BMDM,");
            strSql.Append("JYBH=@JYBH,");
            strSql.Append("SJ=@SJ,");
            strSql.Append("SFZMHM=@SFZMHM,");
            strSql.Append("JSID=@JSID,");
            strSql.Append("JYLX=@JYLX,");
            strSql.Append("GXSJ=@GXSJ,");
            strSql.Append("LDJB=@LDJB");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
	          		new SqlParameter("@XM", SqlDbType.VarChar,50),                                      
                    new SqlParameter("@BMDM", SqlDbType.VarChar,50),                                
                    new SqlParameter("@JYBH", SqlDbType.VarChar,50),
                    new SqlParameter("@SJ", SqlDbType.VarChar,50),    
    				new SqlParameter("@SFZMHM", SqlDbType.VarChar,50),
                    new SqlParameter("@JSID", SqlDbType.Int,4),
                    new SqlParameter("@JYLX", SqlDbType.Int,4),
                     new SqlParameter("@GXSJ", SqlDbType.DateTime),
                    new SqlParameter("@LDJB", SqlDbType.VarChar,2),
                    new SqlParameter("@ID", SqlDbType.BigInt,8)
               
				};
            parameters[0].Value = model.XM;
            parameters[1].Value = model.BMDM;
            parameters[2].Value = model.JYBH;
            parameters[3].Value = model.SJ;
            parameters[4].Value = model.SFZMHM;
            parameters[5].Value = model.JSID;
            parameters[6].Value = model.JYLX;
            parameters[7].Value = model.GXSJ;
            parameters[8].Value = model.LDJB;
            parameters[9].Value = model.ID;
       
            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;

        }




        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_ACL_USER model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update ACL_USER set ");
            strSql.Append("JYBH=@JYBH,");
            strSql.Append("XM=@XM,");
            strSql.Append("BMDM=@BMDM,");
            strSql.Append("SFZMHM=@SFZMHM,");
            strSql.Append("JTZZ=@JTZZ,");
            strSql.Append("JYLX=@JYLX,");
            strSql.Append("MM=@MM,");
            strSql.Append("JSID=@JSID,");
            strSql.Append("SJ=@SJ,");
            strSql.Append("BGDH=@BGDH,");
            strSql.Append("BZLX=@BZLX,");
            strSql.Append("YWGW=@YWGW,");
            strSql.Append("SGCLDJ=@SGCLDJ,");
            strSql.Append("LDJB=@LDJB,");
            strSql.Append("CSRQ=@CSRQ,");
            strSql.Append("XB=@XB,");
            strSql.Append("JX=@JX,");
            strSql.Append("JB=@JB,");
            strSql.Append("JG=@JG,");
            strSql.Append("RDTSJ=@RDTSJ,");
            strSql.Append("ZZMM=@ZZMM,");
            strSql.Append("MZ=@MZ,");
            strSql.Append("XL=@XL,");
            strSql.Append("ZY=@ZY,");
            strSql.Append("ZW=@ZW,");
            strSql.Append("RDSJ=@RDSJ,");
            strSql.Append("CGSJ=@CGSJ,");
            strSql.Append("RXZSJ=@RXZSJ,");
            strSql.Append("ZFZGDJ=@ZFZGDJ,");
            strSql.Append("ZT=@ZT,");
            strSql.Append("GXSJ=@GXSJ,");
            strSql.Append("CJSJ=@CJSJ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@JYBH", SqlDbType.VarChar,50),
					new SqlParameter("@XM", SqlDbType.VarChar,50),
					new SqlParameter("@BMDM", SqlDbType.VarChar,50),
					new SqlParameter("@SFZMHM", SqlDbType.VarChar,50),
					new SqlParameter("@JTZZ", SqlDbType.VarChar,150),
					new SqlParameter("@JYLX", SqlDbType.Int,4),
					new SqlParameter("@MM", SqlDbType.VarChar,50),
					new SqlParameter("@JSID", SqlDbType.Int,4),
					new SqlParameter("@SJ", SqlDbType.VarChar,50),
					new SqlParameter("@BGDH", SqlDbType.VarChar,50),
					new SqlParameter("@BZLX", SqlDbType.VarChar,1),
					new SqlParameter("@YWGW", SqlDbType.VarChar,2),
					new SqlParameter("@SGCLDJ", SqlDbType.VarChar,2),
					new SqlParameter("@LDJB", SqlDbType.VarChar,2),
					new SqlParameter("@CSRQ", SqlDbType.DateTime),
					new SqlParameter("@XB", SqlDbType.VarChar,1),
					new SqlParameter("@JX", SqlDbType.VarChar,10),
					new SqlParameter("@JB", SqlDbType.VarChar,50),
					new SqlParameter("@JG", SqlDbType.VarChar,50),
					new SqlParameter("@RDTSJ", SqlDbType.DateTime),
					new SqlParameter("@ZZMM", SqlDbType.VarChar,50),
					new SqlParameter("@MZ", SqlDbType.VarChar,50),
					new SqlParameter("@XL", SqlDbType.VarChar,50),
					new SqlParameter("@ZY", SqlDbType.VarChar,50),
					new SqlParameter("@ZW", SqlDbType.VarChar,50),
					new SqlParameter("@RDSJ", SqlDbType.DateTime),
					new SqlParameter("@CGSJ", SqlDbType.DateTime),
					new SqlParameter("@RXZSJ", SqlDbType.DateTime),
					new SqlParameter("@ZFZGDJ", SqlDbType.VarChar,50),
					new SqlParameter("@ZT", SqlDbType.Int,4),
					new SqlParameter("@GXSJ", SqlDbType.DateTime),
					new SqlParameter("@CJSJ", SqlDbType.DateTime),
					new SqlParameter("@ID", SqlDbType.BigInt,8)};
            parameters[0].Value = model.JYBH;
            parameters[1].Value = model.XM;
            parameters[2].Value = model.BMDM;
            parameters[3].Value = model.SFZMHM;
            parameters[4].Value = model.JTZZ;
            parameters[5].Value = model.JYLX;
            parameters[6].Value = model.MM;
            parameters[7].Value = model.JSID;
            parameters[8].Value = model.SJ;
            parameters[9].Value = model.BGDH;
            parameters[10].Value = model.BZLX;
            parameters[11].Value = model.YWGW;
            parameters[12].Value = model.SGCLDJ;
            parameters[13].Value = model.LDJB;
            parameters[14].Value = model.CSRQ;
            parameters[15].Value = model.XB;
            parameters[16].Value = model.JX;
            parameters[17].Value = model.JB;
            parameters[18].Value = model.JG;
            parameters[19].Value = model.RDTSJ;
            parameters[20].Value = model.ZZMM;
            parameters[21].Value = model.MZ;
            parameters[22].Value = model.XL;
            parameters[23].Value = model.ZY;
            parameters[24].Value = model.ZW;
            parameters[25].Value = model.RDSJ;
            parameters[26].Value = model.CGSJ;
            parameters[27].Value = model.RXZSJ;
            parameters[28].Value = model.ZFZGDJ;
            parameters[29].Value = model.ZT;
            parameters[30].Value = model.GXSJ;
            parameters[31].Value = model.CJSJ;
            parameters[32].Value = model.ID;


            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;

        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(long ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from ACL_USER ");
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
        /// 删除一条数据根据警员编号来
        /// </summary>
        public bool DeleteJYBH(string JYBH)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from ACL_USER ");
            strSql.Append(" where JYBH=@JYBH");
            SqlParameter[] parameters = {
					new SqlParameter("@JYBH", SqlDbType.VarChar,50),
			};
            parameters[0].Value = JYBH;

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
        /// 得到一部分对象实体
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public Model.M_ACL_USER GetLittleModel(string JYBH)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("SELECT ID, XM,BMDM,JYBH,SJ FROM ACL_USER");
            strSql.Append(" where JYBH=@JYBH");
            SqlParameter[] parameters = {
                 	new SqlParameter("@JYBH", SqlDbType.VarChar,50),
            };
            parameters[0].Value = JYBH;

              Model.M_ACL_USER model = new Model.M_ACL_USER();
              using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
              {

                  while (read.Read())
                  {
                      if (read["ID"] != null && read["ID"].ToString() != "")
                      {
                          model.ID = long.Parse(read["ID"].ToString());
                      }

                      if (read["XM"] != null)
                      {
                          model.XM = read["XM"].ToString();

                      }
                      if (read["BMDM"] != null)
                      {
                          model.BMDM = read["BMDM"].ToString();
                      }

                      if (read["JYBH"] != null)
                      {
                          model.JYBH = read["JYBH"].ToString();
                      }
                      if (read["SJ"] != null)
                      {
                          model.SJ = read["SJ"].ToString();
                      }


                  }
              }

              return model;

        
        }


         
        /// <summary>
        /// 得到一个对象实体
        /// </summary>
        public Model.M_ACL_USER GetModel(long ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,JYBH,XM,BMDM,SFZMHM,JTZZ,JYLX,MM,JSID,SJ,BGDH,BZLX,YWGW,SGCLDJ,LDJB,CSRQ,XB,JX,JB,JG,RDTSJ,ZZMM,MZ,XL,ZY,ZW,RDSJ,CGSJ,RXZSJ,ZFZGDJ,ZT,GXSJ,CJSJ from ACL_USER ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
                    new SqlParameter("@ID", SqlDbType.BigInt)
            };
            parameters[0].Value = ID;

            Model.M_ACL_USER model = new Model.M_ACL_USER();
          using ( SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
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
                if (read["XM"] != null)
                {
                    model.XM = read["XM"].ToString();
                }
                if (read["BMDM"] != null)
                {
                    model.BMDM = read["BMDM"].ToString();
                }
                if (read["SFZMHM"] != null)
                {
                    model.SFZMHM = read["SFZMHM"].ToString();
                }
                if (read["JTZZ"] != null)
                {
                    model.JTZZ = read["JTZZ"].ToString();
                }
                if (read["JYLX"] != null && read["JYLX"].ToString() != "")
                {
                    model.JYLX = int.Parse(read["JYLX"].ToString());
                }
                if (read["MM"] != null)
                {
                    model.MM = read["MM"].ToString();
                }
                if (read["JSID"] != null && read["JSID"].ToString() != "")
                {
                    model.JSID = int.Parse(read["JSID"].ToString());
                }
                if (read["SJ"] != null)
                {
                    model.SJ = read["SJ"].ToString();
                }
                if (read["BGDH"] != null)
                {
                    model.BGDH = read["BGDH"].ToString();
                }
                if (read["BZLX"] != null)
                {
                    model.BZLX = read["BZLX"].ToString();
                }
                if (read["YWGW"] != null)
                {
                    model.YWGW = read["YWGW"].ToString();
                }
                if (read["SGCLDJ"] != null)
                {
                    model.SGCLDJ = read["SGCLDJ"].ToString();
                }
                if (read["LDJB"] != null)
                {
                    model.LDJB = read["LDJB"].ToString();
                }
                if (read["CSRQ"] != null && read["CSRQ"].ToString() != "")
                {
                    model.CSRQ = DateTime.Parse(read["CSRQ"].ToString());
                }
                if (read["XB"] != null)
                {
                    model.XB = read["XB"].ToString();
                }
                if (read["JX"] != null)
                {
                    model.JX = read["JX"].ToString();
                }
                if (read["JB"] != null)
                {
                    model.JB = read["JB"].ToString();
                }
                if (read["JG"] != null)
                {
                    model.JG = read["JG"].ToString();
                }
                if (read["RDTSJ"] != null && read["RDTSJ"].ToString() != "")
                {
                    model.RDTSJ = DateTime.Parse(read["RDTSJ"].ToString());
                }
                if (read["ZZMM"] != null)
                {
                    model.ZZMM = read["ZZMM"].ToString();
                }
                if (read["MZ"] != null)
                {
                    model.MZ = read["MZ"].ToString();
                }
                if (read["XL"] != null)
                {
                    model.XL = read["XL"].ToString();
                }
                if (read["ZY"] != null)
                {
                    model.ZY = read["ZY"].ToString();
                }
                if (read["ZW"] != null)
                {
                    model.ZW = read["ZW"].ToString();
                }
                if (read["RDSJ"] != null && read["RDSJ"].ToString() != "")
                {
                    model.RDSJ = DateTime.Parse(read["RDSJ"].ToString());
                }
                if (read["CGSJ"] != null && read["CGSJ"].ToString() != "")
                {
                    model.CGSJ = DateTime.Parse(read["CGSJ"].ToString());
                }
                if (read["RXZSJ"] != null && read["RXZSJ"].ToString() != "")
                {
                    model.RXZSJ = DateTime.Parse(read["RXZSJ"].ToString());
                }
                if (read["ZFZGDJ"] != null)
                {
                    model.ZFZGDJ = read["ZFZGDJ"].ToString();
                }
                if (read["ZT"] != null && read["ZT"].ToString() != "")
                {
                    model.ZT = int.Parse(read["ZT"].ToString());
                }
                if (read["GXSJ"] != null && read["GXSJ"].ToString() != "")
                {
                    model.GXSJ = DateTime.Parse(read["GXSJ"].ToString());
                }
                if (read["CJSJ"] != null && read["CJSJ"].ToString() != "")
                {
                    model.CJSJ = DateTime.Parse(read["CJSJ"].ToString());
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

            string strSql = "select * from ACL_USER where " + field + "like'" + fieldvalue + "%'";

            DataTable ds = sqlhelper.ExecuteTable(strSql, CommandType.Text);

            return ds;
        }



        /// <summary>
        /// 获得所有列表
        /// </summary>
        /// <returns></returns>
        public  DataTable GetList()
        {
            string sql = "select * from ACL_USER";
            using (var reader =  sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }


        /// <summary>
        /// 更具 角色ID 查找角色是否存在 JSID
        /// </summary>
        /// <returns></returns>
        public bool ExistsJSID(int JSID)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select count(1) from ACL_USER");
            strSql.Append(" where JSID=@JSID");
            SqlParameter[] parameters = {
			new SqlParameter("@JSID", SqlDbType.Int,4),
			};
            parameters[0].Value = JSID;


            return sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text, parameters) > 0;
        }



        /// <summary>
        /// 获取警员总数量
        /// </summary>
        /// <returns></returns>
        public int getcount(string three, string second, int JYLX, int JSID, int LDJB, string search)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select COUNT(XM)  from ACL_USER as A");
            strSql.Append(" left join  PoliceType as B on  A.JYLX=B.ID");
            strSql.Append(" left join Role as C on A.JSID=C.ID  ");

            strSql.Append("  left join Entity as D on A.BMDM=D.BMDM ");
            strSql.Append(" left join  Position as E on A.LDJB=E.ID where 1=1");

            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and A.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (A.BMDM='" + second + "'");

                strSql.Append("  or D.SJBM='" + second + "')");//上级部门编号
            }
            if (JYLX != -2)
            {

                strSql.Append("  and  A.JYLX=" + JYLX);
            }

            if (JSID != -2)
            {
                strSql.Append("  and  A.JSID=" + JSID);

            }

            if (LDJB != -2)
            {
                strSql.Append(" and A.LDJB=" + LDJB);

            }


            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append("  and  A.JYBH='" + search + "'");

            }


            return int.Parse(sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text).ToString());
        }


        /// <summary>
        /// 分页排序警员信息
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(string three, string second, int JYLX, int JSID, int LDJB,string search, int startRowIndex, int maximumRows)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select *  from ACL_USER as A");
            strSql.Append(" left join  PoliceType as B on  A.JYLX=B.ID");
            strSql.Append(" left join Role as C on A.JSID=C.ID ");
            strSql.Append("  left join Entity as D on A.BMDM=D.BMDM ");
            strSql.Append(" left join  Position as E on A.LDJB=E.ID where 1=1");
      


            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and A.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (A.BMDM='" + second + "'");

                strSql.Append("  or D.SJBM='" + second + "')");//上级部门编号
            }


            if (JYLX != -2)
            {

                strSql.Append("  and  A.JYLX=" + JYLX);
            }

            if (JSID != -2)
            {
                strSql.Append("  and  A.JSID=" + JSID);

            }

            if (LDJB != -2)
            {
                strSql.Append(" and A.LDJB=" + LDJB);
            
            }




            if (!string.IsNullOrEmpty(search))
            {

                strSql.Append("  and  A.JYBH='" + search+"'");

            }



            strSql.Append("   order by E.Weight DESC,A.JYBH  DESC  ");

            SqlParameter[] parameters = { };
            return sqlhelper.ExecuteRead(CommandType.Text, strSql.ToString(), startRowIndex, maximumRows, "device", parameters);

        }






        #endregion  BasicMethod
    
    }
}
