using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace JingWuTong.Handle
{
    /// <summary>
    /// OperationLog 的摘要说明
    /// </summary>
    public class OperationLog : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";

            try
            {
                //添加操作日志
                Model.M_OperationLog M_OperationLog_Add = new Model.M_OperationLog();
                BLL.B_OperationLog B_OperationLog_Add = new BLL.B_OperationLog();


                if (HttpContext.Current.Request.Cookies["cookieName"] != null)
                {
                    HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];
                    M_OperationLog_Add.JYBH = cookies["JYBH"];
                }
                M_OperationLog_Add.Module = "00";
                M_OperationLog_Add.OperContent = "02";
                M_OperationLog_Add.LogTime = System.DateTime.Now;

              
                context.Response.Write( B_OperationLog_Add.Add(M_OperationLog_Add));
             

            }

            catch (Exception ex)
            {
               

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