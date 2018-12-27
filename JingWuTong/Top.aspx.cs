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
            if (!Page.IsPostBack)
            {

                if (HttpContext.Current.Request.Cookies["cookieName"] != null)
                {
                    HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];
                    this.LoginName.InnerText = Server.UrlDecode(cookies["JYBH"]);

                }



                BLL.B_Role B_Role_Load = new BLL.B_Role();

                string[] s_power = ((string)B_Role_Load.Exists(TOOL.Login.I_JSID)).Split('|');//权限参数
              

                StringBuilder strjs = new StringBuilder();



                if (s_power[0] == "0")
                {

                    strjs.Append(" $('ul li:eq(0)').hide();"); //首页 

                }


                if (s_power[2] == "0")
                {

                    strjs.Append(" $('ul li:eq(1)').hide();"); //设备查看

                }


                if (s_power[4] == "0")
                {

                    strjs.Append(" $('ul li:eq(2)').hide();"); //实时状况

                }

                if (s_power[6] == "0")
                {

                    strjs.Append(" $('ul li:eq(3)').hide();"); //数据统计

                }

                if (s_power[10] == "0")
                {

                    strjs.Append(" $('ul li:eq(4)').hide();"); //设备管理

                }


                if (s_power[18] == "0")
                {

                    strjs.Append(" $('ul li:eq(5)').hide();"); //人员管理

                }




                if (s_power[24] == "0")
                {

                    strjs.Append(" $('ul li:eq(6)').hide();"); //系统设置

                }

           
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>" + strjs.ToString() + "</script>");

            }
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