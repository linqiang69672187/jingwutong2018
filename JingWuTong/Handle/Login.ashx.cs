using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace JingWuTong.Handle
{
    /// <summary>
    /// Login 的摘要说明
    /// </summary>
    public class Login : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";

       TOOL.Login.strWidth=context.Request.Form["strWidth"];

       TOOL.Login.strHeight = context.Request.Form["strHeight"];

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