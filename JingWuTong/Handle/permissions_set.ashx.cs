using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace JingWuTong.Handle
{
    /// <summary>
    /// permissions_set 的摘要说明
    /// </summary>
    public class permissions_set : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            string type = "新增";

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