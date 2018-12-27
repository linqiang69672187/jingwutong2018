using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace JingWuTong
{
    public class OperationLog : System.Web.UI.Page
    {
     

        /// <summary>
        ///添加操作日志
        /// </summary>
        /// <param name="s_JYBH">警员编号</param>
        /// <param name="s_Module">模块ID</param>
        /// <param name="s_OperContent">操作ID</param>
        /// <returns></returns>
        public  string OperationLogAdd( string s_Module, string s_OperContent)
        {

            try
            {
                string s_JYBH="";
                Model.M_OperationLog M_OperationLog_Add = new Model.M_OperationLog();
                BLL.B_OperationLog B_OperationLog_Add = new BLL.B_OperationLog();

                if (HttpContext.Current.Request.Cookies["cookieName"] != null)
                {
                    HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];
                 s_JYBH = Server.UrlDecode(cookies["JYBH"]);

                }


                M_OperationLog_Add.JYBH = s_JYBH;
                M_OperationLog_Add.Module = s_Module;
                M_OperationLog_Add.OperContent = s_OperContent;
                M_OperationLog_Add.LogTime = System.DateTime.Now;
                M_OperationLog_Add.IpAddr = TOOL.Login.IP;
                M_OperationLog_Add.BZ = TOOL.Login.BZ;
                M_OperationLog_Add.OptObject = TOOL.Login.OptObject;

                B_OperationLog_Add.Add(M_OperationLog_Add);

                return "success ";
            }

            catch (Exception ex)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('" + ex.Message + "');</script>");
                return ex.Message;

            }


           
        
        }

    }
}