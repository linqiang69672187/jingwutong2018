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
    /// 数据访问类:Entity
    /// </summary>
    public partial class Entity
    {
        public Entity()
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
            strSql.Append("select count(1) from Entity");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)
			};
            parameters[0].Value = ID;


            return sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text, parameters) > 0;
        }



        /// <summary>
        /// 根据部门代码查找是否存在
        /// </summary>
        public bool ExistsBMDM(string BMDM)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select count(1) from Entity");
            strSql.Append(" where BMDM=@BMDM");
            SqlParameter[] parameters = {
					new SqlParameter("@BMDM", SqlDbType.VarChar,50)
			};
            parameters[0].Value = BMDM;


            return sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text, parameters) > 0;
        }



        /// <summary>
        /// 根据部门代码查找sjbm
        /// </summary>
        public object ExistsSJBM(string BMDM)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select SJBM  from Entity");
            strSql.Append(" where BMDM=@BMDM");
            SqlParameter[] parameters = {
					new SqlParameter("@BMDM", SqlDbType.VarChar,50)
			};
            parameters[0].Value = BMDM;


            return sqlhelper.ExecuteScalarOne(strSql.ToString(), CommandType.Text, parameters);
        }



        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.Entity model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into Entity(");
            strSql.Append("BMMC,BMQC,BMJC,BMDM,SJBM,Depth,PicUrl,UserCount,YZMC,FZJG,BMJB,KCLYW,YWLB,CZHM,LXDZ,TXZQFR,FZR,LXR,LXDH,JKYH,JZJB,GLTZ,JFLY,YFLY,JLZT,JRGAW,LSGX,JZPTGLBM,CSBJ,FY,FYJG,Sort,Lo,La,BZ,CJSJ,GXSJ)");
            strSql.Append(" values (");
            strSql.Append("@BMMC,@BMQC,@BMJC,@BMDM,@SJBM,@Depth,@PicUrl,@UserCount,@YZMC,@FZJG,@BMJB,@KCLYW,@YWLB,@CZHM,@LXDZ,@TXZQFR,@FZR,@LXR,@LXDH,@JKYH,@JZJB,@GLTZ,@JFLY,@YFLY,@JLZT,@JRGAW,@LSGX,@JZPTGLBM,@CSBJ,@FY,@FYJG,@Sort,@Lo,@La,@BZ,@CJSJ,@GXSJ)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
					new SqlParameter("@BMMC", SqlDbType.VarChar,100),
					new SqlParameter("@BMQC", SqlDbType.VarChar,100),
					new SqlParameter("@BMJC", SqlDbType.VarChar,50),
					new SqlParameter("@BMDM", SqlDbType.VarChar,50),
					new SqlParameter("@SJBM", SqlDbType.VarChar,50),
					new SqlParameter("@Depth", SqlDbType.VarChar,3),
					new SqlParameter("@PicUrl", SqlDbType.VarChar,100),
					new SqlParameter("@UserCount", SqlDbType.Int,4),
					new SqlParameter("@YZMC", SqlDbType.VarChar,50),
					new SqlParameter("@FZJG", SqlDbType.VarChar,50),
					new SqlParameter("@BMJB", SqlDbType.VarChar,1),
					new SqlParameter("@KCLYW", SqlDbType.VarChar,10),
					new SqlParameter("@YWLB", SqlDbType.VarChar,20),
					new SqlParameter("@CZHM", SqlDbType.VarChar,50),
					new SqlParameter("@LXDZ", SqlDbType.VarChar,100),
					new SqlParameter("@TXZQFR", SqlDbType.VarChar,50),
					new SqlParameter("@FZR", SqlDbType.VarChar,50),
					new SqlParameter("@LXR", SqlDbType.VarChar,50),
					new SqlParameter("@LXDH", SqlDbType.VarChar,50),
					new SqlParameter("@JKYH", SqlDbType.VarChar,50),
					new SqlParameter("@JZJB", SqlDbType.VarChar,2),
					new SqlParameter("@GLTZ", SqlDbType.VarChar,1),
					new SqlParameter("@JFLY", SqlDbType.VarChar,1),
					new SqlParameter("@YFLY", SqlDbType.VarChar,1),
					new SqlParameter("@JLZT", SqlDbType.VarChar,1),
					new SqlParameter("@JRGAW", SqlDbType.VarChar,1),
					new SqlParameter("@LSGX", SqlDbType.VarChar,1),
					new SqlParameter("@JZPTGLBM", SqlDbType.VarChar,14),
					new SqlParameter("@CSBJ", SqlDbType.VarChar,1),
					new SqlParameter("@FY", SqlDbType.VarChar,50),
					new SqlParameter("@FYJG", SqlDbType.VarChar,50),
					new SqlParameter("@Sort", SqlDbType.Int,4),
					new SqlParameter("@Lo", SqlDbType.Float,8),
					new SqlParameter("@La", SqlDbType.Float,8),
					new SqlParameter("@BZ", SqlDbType.VarChar,100),
					new SqlParameter("@CJSJ", SqlDbType.DateTime),
					new SqlParameter("@GXSJ", SqlDbType.DateTime)};
            parameters[0].Value = model.BMMC;
            parameters[1].Value = model.BMQC;
            parameters[2].Value = model.BMJC;
            parameters[3].Value = model.BMDM;
            parameters[4].Value = model.SJBM;
            parameters[5].Value = model.Depth;
            parameters[6].Value = model.PicUrl;
            parameters[7].Value = model.UserCount;
            parameters[8].Value = model.YZMC;
            parameters[9].Value = model.FZJG;
            parameters[10].Value = model.BMJB;
            parameters[11].Value = model.KCLYW;
            parameters[12].Value = model.YWLB;
            parameters[13].Value = model.CZHM;
            parameters[14].Value = model.LXDZ;
            parameters[15].Value = model.TXZQFR;
            parameters[16].Value = model.FZR;
            parameters[17].Value = model.LXR;
            parameters[18].Value = model.LXDH;
            parameters[19].Value = model.JKYH;
            parameters[20].Value = model.JZJB;
            parameters[21].Value = model.GLTZ;
            parameters[22].Value = model.JFLY;
            parameters[23].Value = model.YFLY;
            parameters[24].Value = model.JLZT;
            parameters[25].Value = model.JRGAW;
            parameters[26].Value = model.LSGX;
            parameters[27].Value = model.JZPTGLBM;
            parameters[28].Value = model.CSBJ;
            parameters[29].Value = model.FY;
            parameters[30].Value = model.FYJG;
            parameters[31].Value = model.Sort;
            parameters[32].Value = model.Lo;
            parameters[33].Value = model.La;
            parameters[34].Value = model.BZ;
            parameters[35].Value = model.CJSJ;
            parameters[36].Value = model.GXSJ;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            return rows;
        }


        /// <summary>
        /// 增加一条数据导入数据
        /// </summary>
        public int AddEnter(Model.Entity model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into Entity(");
            strSql.Append("BMMC,BMQC,BMJC,BMDM,SJBM,Depth,PicUrl,UserCount,YZMC,FZJG,BMJB,KCLYW,YWLB,CZHM,LXDZ,TXZQFR,FZR,LXR,LXDH,JKYH,JZJB,GLTZ,JFLY,YFLY,JLZT,JRGAW,LSGX,JZPTGLBM,CSBJ,FY,FYJG,Sort,Lo,La,BZ,CJSJ,GXSJ)");
            strSql.Append(" values (");
            strSql.Append("@BMMC,@BMQC,@BMJC,@BMDM,@SJBM,@Depth,@PicUrl,@UserCount,@YZMC,@FZJG,@BMJB,@KCLYW,@YWLB,@CZHM,@LXDZ,@TXZQFR,@FZR,@LXR,@LXDH,@JKYH,@JZJB,@GLTZ,@JFLY,@YFLY,@JLZT,@JRGAW,@LSGX,@JZPTGLBM,@CSBJ,@FY,@FYJG,@Sort,@Lo,@La,@BZ,@CJSJ,@GXSJ)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
					new SqlParameter("@BMMC", SqlDbType.VarChar,100),
					new SqlParameter("@BMQC", SqlDbType.VarChar,100),
					new SqlParameter("@BMJC", SqlDbType.VarChar,50),
					new SqlParameter("@BMDM", SqlDbType.VarChar,50),
					new SqlParameter("@SJBM", SqlDbType.VarChar,50),
					new SqlParameter("@Depth", SqlDbType.VarChar,3),
					new SqlParameter("@PicUrl", SqlDbType.VarChar,100),
					new SqlParameter("@UserCount", SqlDbType.Int,4),
					new SqlParameter("@YZMC", SqlDbType.VarChar,50),
					new SqlParameter("@FZJG", SqlDbType.VarChar,50),
					new SqlParameter("@BMJB", SqlDbType.VarChar,1),
					new SqlParameter("@KCLYW", SqlDbType.VarChar,10),
					new SqlParameter("@YWLB", SqlDbType.VarChar,20),
					new SqlParameter("@CZHM", SqlDbType.VarChar,50),
					new SqlParameter("@LXDZ", SqlDbType.VarChar,100),
					new SqlParameter("@TXZQFR", SqlDbType.VarChar,50),
					new SqlParameter("@FZR", SqlDbType.VarChar,50),
					new SqlParameter("@LXR", SqlDbType.VarChar,50),
					new SqlParameter("@LXDH", SqlDbType.VarChar,50),
					new SqlParameter("@JKYH", SqlDbType.VarChar,50),
					new SqlParameter("@JZJB", SqlDbType.VarChar,2),
					new SqlParameter("@GLTZ", SqlDbType.VarChar,1),
					new SqlParameter("@JFLY", SqlDbType.VarChar,1),
					new SqlParameter("@YFLY", SqlDbType.VarChar,1),
					new SqlParameter("@JLZT", SqlDbType.VarChar,1),
					new SqlParameter("@JRGAW", SqlDbType.VarChar,1),
					new SqlParameter("@LSGX", SqlDbType.VarChar,1),
					new SqlParameter("@JZPTGLBM", SqlDbType.VarChar,14),
					new SqlParameter("@CSBJ", SqlDbType.VarChar,1),
					new SqlParameter("@FY", SqlDbType.VarChar,50),
					new SqlParameter("@FYJG", SqlDbType.VarChar,50),
					new SqlParameter("@Sort", SqlDbType.Int,4),
					new SqlParameter("@Lo", SqlDbType.Float,8),
					new SqlParameter("@La", SqlDbType.Float,8),
					new SqlParameter("@BZ", SqlDbType.VarChar,100),
					new SqlParameter("@CJSJ", SqlDbType.DateTime),
					new SqlParameter("@GXSJ", SqlDbType.DateTime)};
            parameters[0].Value = Transform.CheckIsNull(model.BMMC);
            parameters[1].Value = Transform.CheckIsNull(model.BMQC);
            parameters[2].Value =Transform.CheckIsNull( model.BMJC);
            parameters[3].Value =Transform.CheckIsNull( model.BMDM);
            parameters[4].Value =Transform.CheckIsNull( model.SJBM);
            parameters[5].Value = Transform.CheckIsNull(model.Depth);
            parameters[6].Value =Transform.CheckIsNull( model.PicUrl);
            parameters[7].Value = Transform.CheckIsNull(model.UserCount);
            parameters[8].Value = Transform.CheckIsNull(model.YZMC);
            parameters[9].Value =Transform.CheckIsNull( model.FZJG);
            parameters[10].Value = Transform.CheckIsNull(model.BMJB);
            parameters[11].Value = Transform.CheckIsNull(model.KCLYW);
            parameters[12].Value =Transform.CheckIsNull( model.YWLB);
            parameters[13].Value = Transform.CheckIsNull(model.CZHM);
            parameters[14].Value = Transform.CheckIsNull(model.LXDZ);
            parameters[15].Value =Transform.CheckIsNull( model.TXZQFR);
            parameters[16].Value = Transform.CheckIsNull(model.FZR);
            parameters[17].Value =Transform.CheckIsNull( model.LXR);
            parameters[18].Value =Transform.CheckIsNull( model.LXDH);
            parameters[19].Value =Transform.CheckIsNull(model.JKYH);
            parameters[20].Value = Transform.CheckIsNull(model.JZJB);
            parameters[21].Value = Transform.CheckIsNull(model.GLTZ);
            parameters[22].Value =Transform.CheckIsNull( model.JFLY);
            parameters[23].Value =Transform.CheckIsNull( model.YFLY);
            parameters[24].Value = Transform.CheckIsNull(model.JLZT);
            parameters[25].Value = Transform.CheckIsNull(model.JRGAW);
            parameters[26].Value = Transform.CheckIsNull(model.LSGX);
            parameters[27].Value = Transform.CheckIsNull(model.JZPTGLBM);
            parameters[28].Value =Transform.CheckIsNull( model.CSBJ);
            parameters[29].Value =Transform.CheckIsNull( model.FY);
            parameters[30].Value = Transform.CheckIsNull(model.FYJG);
            parameters[31].Value = Transform.CheckIsNull(model.Sort);
            parameters[32].Value = Transform.CheckIsNull(model.Lo);
            parameters[33].Value =Transform.CheckIsNull( model.La);
            parameters[34].Value = Transform.CheckIsNull(model.BZ);
            parameters[35].Value = Transform.CheckIsNull(model.CJSJ);
            parameters[36].Value =Transform.CheckIsNull(model.GXSJ);

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);
            return rows;
        }

        /// <summary>
        /// 更新一条导入数据
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public int UpdateEnter(Model.Entity model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Entity set ");
            strSql.Append("BMMC=@BMMC,");
            strSql.Append("BMQC=@BMQC,");
            strSql.Append("BMJC=@BMJC,");
            //strSql.Append("BMDM=@BMDM,");
            strSql.Append("SJBM=@SJBM,");
            strSql.Append("Depth=@Depth,");
            strSql.Append("PicUrl=@PicUrl,");
            strSql.Append("UserCount=@UserCount,");
            strSql.Append("YZMC=@YZMC,");
            strSql.Append("FZJG=@FZJG,");
            strSql.Append("BMJB=@BMJB,");
            strSql.Append("KCLYW=@KCLYW,");
            strSql.Append("YWLB=@YWLB,");
            strSql.Append("CZHM=@CZHM,");
            strSql.Append("LXDZ=@LXDZ,");
            strSql.Append("TXZQFR=@TXZQFR,");
            strSql.Append("FZR=@FZR,");
            strSql.Append("LXR=@LXR,");
            strSql.Append("LXDH=@LXDH,");
            strSql.Append("JKYH=@JKYH,");
            strSql.Append("JZJB=@JZJB,");
            strSql.Append("GLTZ=@GLTZ,");
            strSql.Append("JFLY=@JFLY,");
            strSql.Append("YFLY=@YFLY,");
            strSql.Append("JLZT=@JLZT,");
            strSql.Append("JRGAW=@JRGAW,");
            strSql.Append("LSGX=@LSGX,");
            strSql.Append("JZPTGLBM=@JZPTGLBM,");
            strSql.Append("CSBJ=@CSBJ,");
            strSql.Append("FY=@FY,");
            strSql.Append("FYJG=@FYJG,");
            strSql.Append("Sort=@Sort,");
            strSql.Append("Lo=@Lo,");
            strSql.Append("La=@La,");
            strSql.Append("BZ=@BZ,");
            strSql.Append("CJSJ=@CJSJ,");
            strSql.Append("GXSJ=@GXSJ ");
            //strSql.Append(" where ID=@ID");
            strSql.Append(" where BMDM=@BMDM");
            SqlParameter[] parameters = {
					new SqlParameter("@BMMC", SqlDbType.VarChar,100),
					new SqlParameter("@BMQC", SqlDbType.VarChar,100),
					new SqlParameter("@BMJC", SqlDbType.VarChar,50),
                    //new SqlParameter("@BMDM", SqlDbType.VarChar,50),
					new SqlParameter("@SJBM", SqlDbType.VarChar,50),
					new SqlParameter("@Depth", SqlDbType.VarChar,3),
					new SqlParameter("@PicUrl", SqlDbType.VarChar,100),
					new SqlParameter("@UserCount", SqlDbType.Int,4),
					new SqlParameter("@YZMC", SqlDbType.VarChar,50),
					new SqlParameter("@FZJG", SqlDbType.VarChar,50),
					new SqlParameter("@BMJB", SqlDbType.VarChar,1),
					new SqlParameter("@KCLYW", SqlDbType.VarChar,10),
					new SqlParameter("@YWLB", SqlDbType.VarChar,20),
					new SqlParameter("@CZHM", SqlDbType.VarChar,50),
					new SqlParameter("@LXDZ", SqlDbType.VarChar,100),
					new SqlParameter("@TXZQFR", SqlDbType.VarChar,50),
					new SqlParameter("@FZR", SqlDbType.VarChar,50),
					new SqlParameter("@LXR", SqlDbType.VarChar,50),
					new SqlParameter("@LXDH", SqlDbType.VarChar,50),
					new SqlParameter("@JKYH", SqlDbType.VarChar,50),
					new SqlParameter("@JZJB", SqlDbType.VarChar,2),
					new SqlParameter("@GLTZ", SqlDbType.VarChar,1),
					new SqlParameter("@JFLY", SqlDbType.VarChar,1),
					new SqlParameter("@YFLY", SqlDbType.VarChar,1),
					new SqlParameter("@JLZT", SqlDbType.VarChar,1),
					new SqlParameter("@JRGAW", SqlDbType.VarChar,1),
					new SqlParameter("@LSGX", SqlDbType.VarChar,1),
					new SqlParameter("@JZPTGLBM", SqlDbType.VarChar,14),
					new SqlParameter("@CSBJ", SqlDbType.VarChar,1),
					new SqlParameter("@FY", SqlDbType.VarChar,50),
					new SqlParameter("@FYJG", SqlDbType.VarChar,50),
					new SqlParameter("@Sort", SqlDbType.Int,4),
					new SqlParameter("@Lo", SqlDbType.Float,8),
					new SqlParameter("@La", SqlDbType.Float,8),
					new SqlParameter("@BZ", SqlDbType.VarChar,100),
					new SqlParameter("@CJSJ", SqlDbType.DateTime),
					new SqlParameter("@GXSJ", SqlDbType.DateTime),
                    //new SqlParameter("@ID", SqlDbType.Int,4)
                   new SqlParameter("@BMDM", SqlDbType.VarChar,50),
                                        
                                        };
            parameters[0].Value =Transform.CheckIsNull( model.BMMC);
            parameters[1].Value = Transform.CheckIsNull( model.BMQC);
            parameters[2].Value = Transform.CheckIsNull( model.BMJC);
            //parameters[3].Value = model.BMDM;
            parameters[3].Value =Transform.CheckIsNull(  model.SJBM);
            parameters[4].Value =Transform.CheckIsNull(  model.Depth);
            parameters[5].Value = Transform.CheckIsNull( model.PicUrl);
            parameters[6].Value = Transform.CheckIsNull( model.UserCount);
            parameters[7].Value =Transform.CheckIsNull(  model.YZMC);
            parameters[8].Value = Transform.CheckIsNull( model.FZJG);
            parameters[9].Value =Transform.CheckIsNull(  model.BMJB);
            parameters[10].Value = Transform.CheckIsNull( model.KCLYW);
            parameters[11].Value =Transform.CheckIsNull(  model.YWLB);
            parameters[12].Value =Transform.CheckIsNull(  model.CZHM);
            parameters[13].Value =Transform.CheckIsNull(  model.LXDZ);
            parameters[14].Value = Transform.CheckIsNull( model.TXZQFR);
            parameters[15].Value = Transform.CheckIsNull( model.FZR);
            parameters[16].Value =Transform.CheckIsNull(  model.LXR);
            parameters[17].Value =Transform.CheckIsNull(  model.LXDH);
            parameters[18].Value = Transform.CheckIsNull( model.JKYH);
            parameters[19].Value = Transform.CheckIsNull( model.JZJB);
            parameters[20].Value =Transform.CheckIsNull(  model.GLTZ);
            parameters[21].Value =Transform.CheckIsNull(  model.JFLY);
            parameters[22].Value =Transform.CheckIsNull(  model.YFLY);
            parameters[23].Value =Transform.CheckIsNull(  model.JLZT);
            parameters[24].Value = Transform.CheckIsNull( model.JRGAW);
            parameters[25].Value =Transform.CheckIsNull(  model.LSGX);
            parameters[26].Value =Transform.CheckIsNull(  model.JZPTGLBM);
            parameters[27].Value = Transform.CheckIsNull( model.CSBJ);
            parameters[28].Value =Transform.CheckIsNull(  model.FY);
            parameters[29].Value = Transform.CheckIsNull( model.FYJG);
            parameters[30].Value =Transform.CheckIsNull(  model.Sort);
            parameters[31].Value =Transform.CheckIsNull(  model.Lo);
            parameters[32].Value =Transform.CheckIsNull(  model.La);
            parameters[33].Value = Transform.CheckIsNull( model.BZ);
            parameters[34].Value = Transform.CheckIsNull( model.CJSJ);
            parameters[35].Value = Transform.CheckIsNull( model.GXSJ);
            //parameters[37].Value = model.ID;
            parameters[36].Value = Transform.CheckIsNull(model.BMDM);

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;

        
        }


        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update( Model.Entity model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Entity set ");
            strSql.Append("BMMC=@BMMC,");
            strSql.Append("BMQC=@BMQC,");
            strSql.Append("BMJC=@BMJC,");
            strSql.Append("BMDM=@BMDM,");
            strSql.Append("SJBM=@SJBM,");
            strSql.Append("Depth=@Depth,");
            strSql.Append("PicUrl=@PicUrl,");
            strSql.Append("UserCount=@UserCount,");
            strSql.Append("YZMC=@YZMC,");
            strSql.Append("FZJG=@FZJG,");
            strSql.Append("BMJB=@BMJB,");
            strSql.Append("KCLYW=@KCLYW,");
            strSql.Append("YWLB=@YWLB,");
            strSql.Append("CZHM=@CZHM,");
            strSql.Append("LXDZ=@LXDZ,");
            strSql.Append("TXZQFR=@TXZQFR,");
            strSql.Append("FZR=@FZR,");
            strSql.Append("LXR=@LXR,");
            strSql.Append("LXDH=@LXDH,");
            strSql.Append("JKYH=@JKYH,");
            strSql.Append("JZJB=@JZJB,");
            strSql.Append("GLTZ=@GLTZ,");
            strSql.Append("JFLY=@JFLY,");
            strSql.Append("YFLY=@YFLY,");
            strSql.Append("JLZT=@JLZT,");
            strSql.Append("JRGAW=@JRGAW,");
            strSql.Append("LSGX=@LSGX,");
            strSql.Append("JZPTGLBM=@JZPTGLBM,");
            strSql.Append("CSBJ=@CSBJ,");
            strSql.Append("FY=@FY,");
            strSql.Append("FYJG=@FYJG,");
            strSql.Append("Sort=@Sort,");
            strSql.Append("Lo=@Lo,");
            strSql.Append("La=@La,");
            strSql.Append("BZ=@BZ,");
            strSql.Append("CJSJ=@CJSJ,");
            strSql.Append("GXSJ=@GXSJ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@BMMC", SqlDbType.VarChar,100),
					new SqlParameter("@BMQC", SqlDbType.VarChar,100),
					new SqlParameter("@BMJC", SqlDbType.VarChar,50),
					new SqlParameter("@BMDM", SqlDbType.VarChar,50),
					new SqlParameter("@SJBM", SqlDbType.VarChar,50),
					new SqlParameter("@Depth", SqlDbType.VarChar,3),
					new SqlParameter("@PicUrl", SqlDbType.VarChar,100),
					new SqlParameter("@UserCount", SqlDbType.Int,4),
					new SqlParameter("@YZMC", SqlDbType.VarChar,50),
					new SqlParameter("@FZJG", SqlDbType.VarChar,50),
					new SqlParameter("@BMJB", SqlDbType.VarChar,1),
					new SqlParameter("@KCLYW", SqlDbType.VarChar,10),
					new SqlParameter("@YWLB", SqlDbType.VarChar,20),
					new SqlParameter("@CZHM", SqlDbType.VarChar,50),
					new SqlParameter("@LXDZ", SqlDbType.VarChar,100),
					new SqlParameter("@TXZQFR", SqlDbType.VarChar,50),
					new SqlParameter("@FZR", SqlDbType.VarChar,50),
					new SqlParameter("@LXR", SqlDbType.VarChar,50),
					new SqlParameter("@LXDH", SqlDbType.VarChar,50),
					new SqlParameter("@JKYH", SqlDbType.VarChar,50),
					new SqlParameter("@JZJB", SqlDbType.VarChar,2),
					new SqlParameter("@GLTZ", SqlDbType.VarChar,1),
					new SqlParameter("@JFLY", SqlDbType.VarChar,1),
					new SqlParameter("@YFLY", SqlDbType.VarChar,1),
					new SqlParameter("@JLZT", SqlDbType.VarChar,1),
					new SqlParameter("@JRGAW", SqlDbType.VarChar,1),
					new SqlParameter("@LSGX", SqlDbType.VarChar,1),
					new SqlParameter("@JZPTGLBM", SqlDbType.VarChar,14),
					new SqlParameter("@CSBJ", SqlDbType.VarChar,1),
					new SqlParameter("@FY", SqlDbType.VarChar,50),
					new SqlParameter("@FYJG", SqlDbType.VarChar,50),
					new SqlParameter("@Sort", SqlDbType.Int,4),
					new SqlParameter("@Lo", SqlDbType.Float,8),
					new SqlParameter("@La", SqlDbType.Float,8),
					new SqlParameter("@BZ", SqlDbType.VarChar,100),
					new SqlParameter("@CJSJ", SqlDbType.DateTime),
					new SqlParameter("@GXSJ", SqlDbType.DateTime),
					new SqlParameter("@ID", SqlDbType.Int,4)};
            parameters[0].Value = model.BMMC;
            parameters[1].Value = model.BMQC;
            parameters[2].Value = model.BMJC;
            parameters[3].Value = model.BMDM;
            parameters[4].Value = model.SJBM;
            parameters[5].Value = model.Depth;
            parameters[6].Value = model.PicUrl;
            parameters[7].Value = model.UserCount;
            parameters[8].Value = model.YZMC;
            parameters[9].Value = model.FZJG;
            parameters[10].Value = model.BMJB;
            parameters[11].Value = model.KCLYW;
            parameters[12].Value = model.YWLB;
            parameters[13].Value = model.CZHM;
            parameters[14].Value = model.LXDZ;
            parameters[15].Value = model.TXZQFR;
            parameters[16].Value = model.FZR;
            parameters[17].Value = model.LXR;
            parameters[18].Value = model.LXDH;
            parameters[19].Value = model.JKYH;
            parameters[20].Value = model.JZJB;
            parameters[21].Value = model.GLTZ;
            parameters[22].Value = model.JFLY;
            parameters[23].Value = model.YFLY;
            parameters[24].Value = model.JLZT;
            parameters[25].Value = model.JRGAW;
            parameters[26].Value = model.LSGX;
            parameters[27].Value = model.JZPTGLBM;
            parameters[28].Value = model.CSBJ;
            parameters[29].Value = model.FY;
            parameters[30].Value = model.FYJG;
            parameters[31].Value = model.Sort;
            parameters[32].Value = model.Lo;
            parameters[33].Value = model.La;
            parameters[34].Value = model.BZ;
            parameters[35].Value = model.CJSJ;
            parameters[36].Value = model.GXSJ;
            parameters[37].Value = model.ID;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;

        }




        /// <summary>
        /// 更新一条数据 部门管理
        /// </summary>
        public int UpdateDepartment(string BMMC, string BMJC, string SJBM, string LXDZ, decimal Lo, decimal La, string BMDM, string FZR, string LXDH, int Sort, string FY, int IsDel, DateTime GXSJ, int ID, string BM, string SJ)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update Entity_Temp set ");

            strSql.Append("BMJC=@BMJC,");

            strSql.Append("LXDZ=@LXDZ,");
            strSql.Append("Lo=@Lo,");
            strSql.Append("La=@La,");
    
            strSql.Append("FZR=@FZR,");
            strSql.Append("LXDH=@LXDH,");
            strSql.Append("Sort=@Sort,");
            strSql.Append("FY=@FY,");

            strSql.Append("IsDel=@IsDel,");
            strSql.Append("GXSJ=@GXSJ ");

            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
		
                    new SqlParameter("@BMJC", SqlDbType.VarChar,50),
		
                    new SqlParameter("@LXDZ", SqlDbType.VarChar,100),
                    new SqlParameter("@Lo", SqlDbType.Float,8),
					new SqlParameter("@La", SqlDbType.Float,8),

					new SqlParameter("@FZR", SqlDbType.VarChar,50),
					new SqlParameter("@LXDH", SqlDbType.VarChar,50),
			        new SqlParameter("@Sort", SqlDbType.Int,4),
                    new SqlParameter("@FY", SqlDbType.VarChar,50),
                     new SqlParameter("@IsDel", SqlDbType.Int,4),//是否显示

					new SqlParameter("@GXSJ", SqlDbType.DateTime),
					new SqlParameter("@ID", SqlDbType.Int,4),
             
                                        };


            parameters[0].Value = Transform.CheckIsNull(BMJC);
    
            parameters[1].Value =Transform.CheckIsNull(LXDZ);
            parameters[2].Value = Transform.CheckIsNull(Lo);
            parameters[3].Value =Transform.CheckIsNull (La);
 
            parameters[4].Value = Transform.CheckIsNull(FZR);
            parameters[5].Value =Transform.CheckIsNull(LXDH);
            parameters[6].Value = Transform.CheckIsNull(Sort); ;
            parameters[7].Value = Transform.CheckIsNull(FY).ToString();
            parameters[8].Value = Transform.CheckIsNull(IsDel);
            parameters[9].Value = GXSJ;
            parameters[10].Value = ID;
     


            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;

        }


        ///// <summary>
        ///// 更新部门是否显示
        ///// </summary>
        ///// <param name="isdel"></param>
        ///// <param name="ID"></param>
        ///// <returns></returns>
        //public int UpdateDepartment_Entity_Temp(int isdel, string BMDM)
        //{

        //    StringBuilder strSql = new StringBuilder();
        //    strSql.Append("update Entity_Temp set ");

        //    strSql.Append("IsDel=@IsDel");

        //    strSql.Append(" where BMDM=@BMDM");
        //    SqlParameter[] parameters = {
		
        //            new SqlParameter("@IsDel", SqlDbType.Int,4),
		
           
        //            new SqlParameter("@BMDM", SqlDbType.VarChar,50)
             
        //                                };


        //    parameters[0].Value = Transform.CheckIsNull(isdel);

        //    parameters[1].Value = Transform.CheckIsNull(BMDM);

        //    int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

        //    return rows;
        
        //}


        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from Entity ");
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
        public Model.Entity GetModel(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,BMMC,BMQC,BMJC,BMDM,SJBM,Depth,PicUrl,UserCount,YZMC,FZJG,BMJB,KCLYW,YWLB,CZHM,LXDZ,TXZQFR,FZR,LXR,LXDH,JKYH,JZJB,GLTZ,JFLY,YFLY,JLZT,JRGAW,LSGX,JZPTGLBM,CSBJ,FY,FYJG,Sort,Lo,La,BZ,CJSJ,GXSJ from Entity ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@ID", SqlDbType.Int,4)
			};
            parameters[0].Value = ID;

           Model.Entity model = new Model.Entity();
           using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text, parameters))
           {

               while (read.Read())
               {
                   if (read["ID"] != null && read["ID"].ToString() != "")
                   {
                       model.ID = int.Parse(read["ID"].ToString());
                   }
                   if (read["BMMC"] != null)
                   {
                       model.BMMC = read["BMMC"].ToString();
                   }
                   if (read["BMQC"] != null)
                   {
                       model.BMQC = read["BMQC"].ToString();
                   }
                   if (read["BMJC"] != null)
                   {
                       model.BMJC = read["BMJC"].ToString();
                   }
                   if (read["BMDM"] != null)
                   {
                       model.BMDM = read["BMDM"].ToString();
                   }
                   if (read["SJBM"] != null)
                   {
                       model.SJBM = read["SJBM"].ToString();
                   }
                   if (read["Depth"] != null)
                   {
                       model.Depth = read["Depth"].ToString();
                   }
                   if (read["PicUrl"] != null)
                   {
                       model.PicUrl = read["PicUrl"].ToString();
                   }
                   if (read["UserCount"] != null && read["UserCount"].ToString() != "")
                   {
                       model.UserCount = int.Parse(read["UserCount"].ToString());
                   }
                   if (read["YZMC"] != null)
                   {
                       model.YZMC = read["YZMC"].ToString();
                   }
                   if (read["FZJG"] != null)
                   {
                       model.FZJG = read["FZJG"].ToString();
                   }
                   if (read["BMJB"] != null)
                   {
                       model.BMJB = read["BMJB"].ToString();
                   }
                   if (read["KCLYW"] != null)
                   {
                       model.KCLYW = read["KCLYW"].ToString();
                   }
                   if (read["YWLB"] != null)
                   {
                       model.YWLB = read["YWLB"].ToString();
                   }
                   if (read["CZHM"] != null)
                   {
                       model.CZHM = read["CZHM"].ToString();
                   }
                   if (read["LXDZ"] != null)
                   {
                       model.LXDZ = read["LXDZ"].ToString();
                   }
                   if (read["TXZQFR"] != null)
                   {
                       model.TXZQFR = read["TXZQFR"].ToString();
                   }
                   if (read["FZR"] != null)
                   {
                       model.FZR = read["FZR"].ToString();
                   }
                   if (read["LXR"] != null)
                   {
                       model.LXR = read["LXR"].ToString();
                   }
                   if (read["LXDH"] != null)
                   {
                       model.LXDH = read["LXDH"].ToString();
                   }
                   if (read["JKYH"] != null)
                   {
                       model.JKYH = read["JKYH"].ToString();
                   }
                   if (read["JZJB"] != null)
                   {
                       model.JZJB = read["JZJB"].ToString();
                   }
                   if (read["GLTZ"] != null)
                   {
                       model.GLTZ = read["GLTZ"].ToString();
                   }
                   if (read["JFLY"] != null)
                   {
                       model.JFLY = read["JFLY"].ToString();
                   }
                   if (read["YFLY"] != null)
                   {
                       model.YFLY = read["YFLY"].ToString();
                   }
                   if (read["JLZT"] != null)
                   {
                       model.JLZT = read["JLZT"].ToString();
                   }
                   if (read["JRGAW"] != null)
                   {
                       model.JRGAW = read["JRGAW"].ToString();
                   }
                   if (read["LSGX"] != null)
                   {
                       model.LSGX = read["LSGX"].ToString();
                   }
                   if (read["JZPTGLBM"] != null)
                   {
                       model.JZPTGLBM = read["JZPTGLBM"].ToString();
                   }
                   if (read["CSBJ"] != null)
                   {
                       model.CSBJ = read["CSBJ"].ToString();
                   }
                   if (read["FY"] != null)
                   {
                       model.FY = read["FY"].ToString();
                   }
                   if (read["FYJG"] != null)
                   {
                       model.FYJG = read["FYJG"].ToString();
                   }
                   if (read["Sort"] != null && read["Sort"].ToString() != "")
                   {
                       model.Sort = int.Parse(read["Sort"].ToString());
                   }
                   if (read["Lo"] != null && read["Lo"].ToString() != "")
                   {
                       model.Lo = decimal.Parse(read["Lo"].ToString());
                   }
                   if (read["La"] != null && read["La"].ToString() != "")
                   {
                       model.La = decimal.Parse(read["La"].ToString());
                   }
                   if (read["BZ"] != null)
                   {
                       model.BZ = read["BZ"].ToString();
                   }
                   if (read["CJSJ"] != null && read["CJSJ"].ToString() != "")
                   {
                       model.CJSJ = DateTime.Parse(read["CJSJ"].ToString());
                   }
                   if (read["GXSJ"] != null && read["GXSJ"].ToString() != "")
                   {
                       model.GXSJ = DateTime.Parse(read["GXSJ"].ToString());
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

            string strSql = "select *  from Entity where " + field + "like'" + fieldvalue + "%'";

            DataTable ds = sqlhelper.ExecuteTable(strSql, CommandType.Text);

            return ds;
        }

        /// <summary>
        /// 获取子单位
        /// </summary>
        /// <param name="SJBM"></param>
        /// <returns></returns>
        public DataTable GetEntity(string SJBM, int DevType, string BMDM)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Append("select BMJC as BMMC,count(DevType) as DevType from Entity as A left join Device as B  on A.BMDM=b.BMDM and B.DevType=@DevType ");
            strSql.Append(" where SJBM=@SJBM and DevType<>0 or A.BMDM=@BMDM  group by A.BMJC  ");
            SqlParameter[] parameters = {
			    new SqlParameter("@SJBM", SqlDbType.VarChar,50),
                new SqlParameter("@DevType", SqlDbType.Int,4),
                new SqlParameter("@BMDM", SqlDbType.VarChar,50)
			};
            parameters[0].Value = SJBM;
            parameters[1].Value = DevType;
            parameters[2].Value = BMDM;

            using (var reader = sqlhelper.ExecuteTable(strSql.ToString(), CommandType.Text, parameters))
            {
                return reader;
            }
        
        }

        /// <summary>
        /// 获取单位终端数
        /// </summary>
        /// <param name="SJBM"></param>
        /// <returns></returns>
        public int GetCount(string SJBM,string BMDM)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append(" select COUNT(B.DevType) as DevType from Entity as A left join Device  as B ");
            strSql.Append("on A.BMDM=B.BMDM where A.SJBM=@SJBM or A.BMDM=@BMDM");
            SqlParameter[] parameters = {
			    new SqlParameter("@SJBM", SqlDbType.VarChar,50),
                new SqlParameter("@BMDM", SqlDbType.VarChar,50)
                  };
            parameters[0].Value = SJBM;
            parameters[1].Value = BMDM;

            int i_count = sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text, parameters);
          
           return i_count;
          

        
        }


        /// <summary>
        /// 获取四大队的设备每个类型总数 
        /// </summary>
        /// <returns></returns>
        public int GetAllDevTyCount(int DevType, string s_BMDM)
        {
            StringBuilder strSql = new StringBuilder();
             strSql.Append("select count(B.DevType) as DevType from Entity as A");
             strSql.Append(" left join Device as B  on A.BMDM=b.BMDM");


             if (s_BMDM == "331000000000")
             {

                 strSql.Append(" where ( A.SJBM in('331002000000','331003000000','331004000000','331001000000') ");
                 strSql.Append(" or A.BMDM in ('331002000000','331003000000','331004000000','331001000000')) and DevType=@DevType");
             }

             else
             {
                 strSql.Append(" where ( A.SJBM ='" + s_BMDM+"'");
                 strSql.Append(" or A.BMDM='" + s_BMDM + "') and DevType=@DevType");
             
             }



            SqlParameter[] parameters = {
            new SqlParameter("@DevType", SqlDbType.Int,4),
                  };
            parameters[0].Value = DevType;


            int i_count = sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text, parameters);

            return i_count;

        }

        /// <summary>
        ///     获取四大队的设备总数
        /// </summary>
        /// <param name="DevType"></param>
        /// <returns></returns>
        public int GetAllCount(string s_BMDM)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select count(B.DevType) as DevType from Entity as A");
            strSql.Append(" left join Device as B  on A.BMDM=b.BMDM");

            if (s_BMDM == "331000000000")
            {
                strSql.Append(" where ( A.SJBM in('331002000000','331003000000','331004000000','331001000000') ");
                strSql.Append(" or A.BMDM in ('331002000000','331003000000','331004000000','331001000000')) ");
            }

            else
            {
                strSql.Append(" where ( A.SJBM='" + s_BMDM+"'");
                strSql.Append(" or A.BMDM ='" + s_BMDM+"')");
            
            }

            int i_count =int.Parse(sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text).ToString());

            return i_count;


        }




        /// <summary>
        /// 获得所有列表
        /// </summary>
        /// <returns></returns>
        public DataTable GetList()
        {
            string sql = "select * from Entity";
            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }




        /// <summary>
        /// 根据部门名称 查找对应BMDM
        /// </summary>
        /// <returns></returns>
        public object GetListName(string BMMC)
        {
            string sql = " select BMDM from Entity where BMMC ='" + BMMC + "' or  BMQC='"+BMMC+"'";

            return sqlhelper.ExecuteScalarOne(sql, CommandType.Text);


        }


        /// <summary>
        /// 根据BMDM查找对应BMDM
        /// </summary>
        /// <returns></returns>
        public object GetListBMMD(string BMDM)
        {
            string sql = " select BMDM from Entity where BMDM  ='" + BMDM + "'";

            return sqlhelper.ExecuteScalarOne(sql, CommandType.Text);


        }



        /// <summary>
        /// 根据部门名称查找部门编号
        /// </summary>
        /// <returns></returns>
        public object GetListBMDM(string s_BMMC)
        {
            string sql = "select BMDM from Entity where BMMC ='" + s_BMMC+"'";
            return sqlhelper.ExecuteScalarOne(sql, CommandType.Text);
        }


        /// <summary>
        /// 获得DropDownList所需字段
        /// </summary>
        /// <returns></returns>
        public DataTable GetList(string s_SJBM, string s_BMDM)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select BMMC,BMDM from Entity ");
            strSql.Append("where SJBM=@SJBM ");

            if (!string.IsNullOrEmpty(s_BMDM))
            {
                strSql.Append("and BMDM=@BMDM ");
            }

            strSql.Append(" order by Sort desc");
            SqlParameter[] parameters = {
				    new SqlParameter("@SJBM", SqlDbType.VarChar,50),
                      new SqlParameter("@BMDM", SqlDbType.VarChar,50)
			};
            parameters[0].Value = s_SJBM;
            parameters[1].Value = s_BMDM;

            using (var reader = sqlhelper.ExecuteTable(strSql.ToString(), CommandType.Text, parameters))
            {
                return reader;
            }
        }




        /// <summary>
        /// 获得DropDownList所需字段
        /// </summary>
        /// <returns></returns>
        public DataTable GetList2(string s_SJBM, string s_BMDM)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select BMMC,BMDM from Entity_Temp ");
            strSql.Append("where SJBM=@SJBM ");

            if (!string.IsNullOrEmpty(s_BMDM))
            {
                strSql.Append("and BMDM=@BMDM ");
            }

            strSql.Append(" order by Sort desc");
            SqlParameter[] parameters = {
				    new SqlParameter("@SJBM", SqlDbType.VarChar,50),
                      new SqlParameter("@BMDM", SqlDbType.VarChar,50)
			};
            parameters[0].Value = s_SJBM;
            parameters[1].Value = s_BMDM;

            using (var reader = sqlhelper.ExecuteTable(strSql.ToString(), CommandType.Text, parameters))
            {
                return reader;
            }
        }





        /// <summary>
        /// 获取设备总数量 部门管理
        /// </summary>
        /// <returns></returns>
        public int getcount(string three, string second, string search)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select count(A.ID) from Entity_Temp as A   ");

            strSql.Append(" left join Entity_Temp as B on A.SJBM=B.BMDM where 1=1  ");


            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and A.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (A.BMDM='" + second + "'");

                strSql.Append("  or A.SJBM='" + second + "')");//上级部门编号
            }


            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and A.BMQC like '%" + search + "%'");

            }


            return int.Parse(sqlhelper.ExecuteScalar(strSql.ToString(), CommandType.Text).ToString());
        }






        /// <summary>
        /// 分页排序设备信息 部门管理
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging( string three,string second, string search, int startRowIndex, int maximumRows)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("select A.ID,A.SJBM,A.BMJC,A.BMMC as BM,B.BMMC as SJ,A.LXDZ,A.Lo,A.La,A.BMDM,A.FZR,A.LXDH,A.JKYH,A.FY,A.FYJG,A.Sort, CASE  A.IsDel WHEN  '0' THEN '是'  WHEN NULL THEN ''  ELSE '否'END  as IsDel1,A.IsDel  from Entity_Temp  as A   ");

            strSql.Append(" left join Entity_Temp  as B on A.SJBM=B.BMDM  ");
            strSql.Append("   where 1=1  ");


            if (!string.IsNullOrEmpty(three))
            {
                strSql.Append(" and A.BMDM= '" + three + "'");
            }

            else if (!string.IsNullOrEmpty(second))
            {
                strSql.Append("  and (A.BMDM='" + second + "'");

                strSql.Append("  or A.SJBM='" + second + "')");//上级部门编号
            }


            if (!string.IsNullOrEmpty(search))
            {
                strSql.Append(" and A.BMQC like '%" + search + "%'");
            }


            strSql.Append(" order by A.Sort desc");

            SqlParameter[] parameters = { };
            return sqlhelper.ExecuteRead(CommandType.Text, strSql.ToString(), startRowIndex, maximumRows, "device", parameters);

        }


        #endregion  BasicMethod


    }
}
