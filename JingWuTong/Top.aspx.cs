using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong
{
    public partial class Top : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
           
               
        }

  


        //protected void B_Out_Click(object sender, EventArgs e)
        //{

        //    try
        //    {
        //        //添加操作日志
        //        Model.M_OperationLog M_OperationLog_Add = new Model.M_OperationLog();
        //        BLL.B_OperationLog B_OperationLog_Add = new BLL.B_OperationLog();


        //        if (HttpContext.Current.Request.Cookies["cookieName"] != null)
        //        {
        //            HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];
        //            M_OperationLog_Add.JYBH = Server.UrlDecode(cookies["JYBH"]);
        //        }
        //        M_OperationLog_Add.Module = "00";
        //        M_OperationLog_Add.OperContent = "02";
        //        M_OperationLog_Add.LogTime = System.DateTime.Now;

        //        B_OperationLog_Add.Add(M_OperationLog_Add);
        //    }

        //    catch (Exception ex)
        //    {
        //        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('" + ex.Message + "');</script>");

        //    }

        //}





    }
}