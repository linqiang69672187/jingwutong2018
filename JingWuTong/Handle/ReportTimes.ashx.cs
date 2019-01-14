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
                foreach (var key in ConfigurationManager.AppSettings.AllKeys)
                {
                    if (!key.Contains("Time")) continue;
                    sb.Append(ConfigurationManager.AppSettings[key] + ",");

                }
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