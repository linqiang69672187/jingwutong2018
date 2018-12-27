using DbComponent;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;

namespace JingWuTong.Handle
{
    /// <summary>
    /// ReportTimes 的摘要说明
    /// </summary>
    public class ReportTimes : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            try
            {
              

        // 获取Config 中的数据
                StringBuilder sb = new StringBuilder();
                sb.Append(ConfigurationManager.AppSettings["time1"]+",");
                sb.Append(ConfigurationManager.AppSettings["time2"] + ",");
                sb.Append(ConfigurationManager.AppSettings["time3"] + ",");
                sb.Append(ConfigurationManager.AppSettings["time4"] + ",");
                sb.Append(ConfigurationManager.AppSettings["time5"] + ",");

                string sql =" select val  from IndexConfigs where  DevType=7 ";

                string s_value = SQLHelper.ExecuteScalar(CommandType.Text, sql.ToString()).ToString();

                sb.Append(s_value);
                context.Response.Write(sb.ToString());


            }

            catch (Exception ex)
            {
               string s= ex.Message;

            }

        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}