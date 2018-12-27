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
    /// 数据访问类:ProjectInfo
    /// </summary>
    public partial class D_ProjectInfo
    {
        public D_ProjectInfo()
        { }
        //实例化
        SqlHelper sqlhelper = new SqlHelper();
        #region  BasicMethod

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_ProjectInfo model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("insert into ProjectInfo(");
            strSql.Append("XMBH,XMMC,XMFZR,XMFZRDH,CGSJ,BXQ,BFQX,Manufacturer,BCJYSJ,XCJYSJ)");
            strSql.Append(" values (");
            strSql.Append("@XMBH,@XMMC,@XMFZR,@XMFZRDH,@CGSJ,@BXQ,@BFQX,@Manufacturer,@BCJYSJ,@XCJYSJ)");
            strSql.Append(";select @@IDENTITY");
            SqlParameter[] parameters = {
					new SqlParameter("@XMBH", SqlDbType.VarChar,50),
					new SqlParameter("@XMMC", SqlDbType.VarChar,100),
					new SqlParameter("@XMFZR", SqlDbType.VarChar,50),
					new SqlParameter("@XMFZRDH", SqlDbType.VarChar,50),
					new SqlParameter("@CGSJ", SqlDbType.Date,3),
					new SqlParameter("@BXQ", SqlDbType.Date,3),
					new SqlParameter("@BFQX", SqlDbType.Date,3),
					new SqlParameter("@Manufacturer", SqlDbType.VarChar,100),
					new SqlParameter("@BCJYSJ", SqlDbType.DateTime),
					new SqlParameter("@XCJYSJ", SqlDbType.DateTime)};
            parameters[0].Value = model.XMBH;
            parameters[1].Value = model.XMMC;
            parameters[2].Value = model.XMFZR;
            parameters[3].Value = model.XMFZRDH;
            parameters[4].Value = model.CGSJ;
            parameters[5].Value = model.BXQ;
            parameters[6].Value = model.BFQX;
            parameters[7].Value = model.Manufacturer;
            parameters[8].Value = model.BCJYSJ;
            parameters[9].Value = model.XCJYSJ;
            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }
        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_ProjectInfo model)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Append("update ProjectInfo set ");
            strSql.Append("XMBH=@XMBH,");
            strSql.Append("XMMC=@XMMC,");
            strSql.Append("XMFZR=@XMFZR,");
            strSql.Append("XMFZRDH=@XMFZRDH,");
            strSql.Append("CGSJ=@CGSJ,");
            strSql.Append("BXQ=@BXQ,");
            strSql.Append("BFQX=@BFQX,");
            strSql.Append("Manufacturer=@Manufacturer,");
            strSql.Append("BCJYSJ=@BCJYSJ,");
            strSql.Append("XCJYSJ=@XCJYSJ");
            strSql.Append(" where ID=@ID");
            SqlParameter[] parameters = {
					new SqlParameter("@XMBH", SqlDbType.VarChar,50),
					new SqlParameter("@XMMC", SqlDbType.VarChar,100),
					new SqlParameter("@XMFZR", SqlDbType.VarChar,50),
					new SqlParameter("@XMFZRDH", SqlDbType.VarChar,50),
					new SqlParameter("@CGSJ", SqlDbType.Date,3),
					new SqlParameter("@BXQ", SqlDbType.Date,3),
					new SqlParameter("@BFQX", SqlDbType.Date,3),
					new SqlParameter("@Manufacturer", SqlDbType.VarChar,100),
					new SqlParameter("@BCJYSJ", SqlDbType.DateTime),
					new SqlParameter("@XCJYSJ", SqlDbType.DateTime),
					new SqlParameter("@ID", SqlDbType.Int,4)};
            parameters[0].Value = model.XMBH;
            parameters[1].Value = model.XMMC;
            parameters[2].Value = model.XMFZR;
            parameters[3].Value = model.XMFZRDH;
            parameters[4].Value = model.CGSJ;
            parameters[5].Value = model.BXQ;
            parameters[6].Value = model.BFQX;
            parameters[7].Value = model.Manufacturer;
            parameters[8].Value = model.BCJYSJ;
            parameters[9].Value = model.XCJYSJ;
            parameters[10].Value = model.ID;

            int rows = sqlhelper.ExecuteNonQuery(strSql.ToString(), CommandType.Text, parameters);

            return rows;
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(int ID)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("delete from ProjectInfo ");
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
        public Model.M_ProjectInfo GetModel(string project)
        {

            StringBuilder strSql = new StringBuilder();
            strSql.Append("select  top 1 ID,XMBH,XMMC,XMFZR,XMFZRDH,CGSJ,BXQ,BFQX,Manufacturer,BCJYSJ,XCJYSJ from ProjectInfo ");
            strSql.Append(" where XMBH='" + project + "'or XMMC='" + project+"'");


            Model.M_ProjectInfo model = new Model.M_ProjectInfo();
            using (SqlDataReader read = sqlhelper.ExecuteSqlReader(strSql.ToString(), CommandType.Text))
            {
                while (read.Read())
                {
                    if (read["ID"] != null && read["ID"].ToString() != "")
                    {
                        model.ID = int.Parse(read["ID"].ToString());
                    }
                    if (read["XMBH"] != null)
                    {
                        model.XMBH = read["XMBH"].ToString();
                    }
                    if (read["XMMC"] != null)
                    {
                        model.XMMC = read["XMMC"].ToString();
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
                    if (read["Manufacturer"] != null)
                    {
                        model.Manufacturer = read["Manufacturer"].ToString();
                    }
                    if (read["BCJYSJ"] != null && read["BCJYSJ"].ToString() != "")
                    {
                        model.BCJYSJ = DateTime.Parse(read["BCJYSJ"].ToString());
                    }
                    if (read["XCJYSJ"] != null && read["XCJYSJ"].ToString() != "")
                    {
                        model.XCJYSJ = DateTime.Parse(read["XCJYSJ"].ToString());
                    }

                }
            }
            return model;
        }






  
        /// <summary>
        /// 获得数据列表
        /// </summary>
        public DataTable GetList()
        {
            string sql = "select * from ProjectInfo";
            using (var reader = sqlhelper.ExecuteTable(sql, CommandType.Text))
            {
                return reader;
            }
        }

        #endregion  BasicMethod

    }
}
