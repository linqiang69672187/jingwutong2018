using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong.Personnel
{
    public partial class Role_Management_Add : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {


                string ID = Request.QueryString["ID"];

                string s_name = "";

                if (HttpContext.Current.Request.Cookies["cookieName"] != null)
                {

                    HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];

                    if (cookies["name"]!=null)
                    { 
                    s_name = Server.UrlDecode(cookies["name"]);
                    }
             

                }



                if (!string.IsNullOrEmpty(ID))
                {
                    this.L_head.InnerText = "修改人员管理";
                    load(int.Parse(ID));

                }
                else
                {

                    this.L_Crateator.InnerText = s_name;
                    this.L_head.InnerText = "新增人员管理";
                
                }





            }

        }





        /// <summary>
        /// 加载
        /// </summary>
        protected void load(int ID)
        {

            //角色表
            BLL.B_Role B_Role_load = new BLL.B_Role();

            Model.M_Role M_Role_load = new Model.M_Role();

           M_Role_load= B_Role_load.GetModel(ID);

           this.T_RoleName.Text = M_Role_load.RoleName;

           BLL.B_ACL_USER B_ACL_USER_load = new BLL.B_ACL_USER();

           this.L_Crateator.InnerText =(string)B_ACL_USER_load.ExistsJYXM( M_Role_load.Crateator);


           this.T_Bz.Text = M_Role_load.Bz;

           this.HiddenPowerLoad.Value = M_Role_load.Power;

           //记入日志
           //TOOL.Login.OptObject = this.T_DevId.Text.Trim();//设备编号
           TOOL.Login.BZ = "操作前:" + M_Role_load.Statistics();//备注

         Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>loadCheck();</script>");


        }






        /// <summary>
        /// 模型公用Role
        /// </summary>
        /// <param name="M_DeviceLog"></param>
        /// <returns></returns>
        protected Model.M_Role All_M_Role(Model.M_Role M_Role)
        {

            M_Role.RoleName = this.T_RoleName.Text.Trim();


            if (HttpContext.Current.Request.Cookies["cookieName"] != null)
            {

                HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];
                if (cookies["JYBH"] != null)
                {
                    M_Role.Crateator = Server.UrlDecode(cookies["JYBH"]);
                }

            }



           

            M_Role.Bz = this.T_Bz.Text.Trim();

            M_Role.Power = this.HiddenPower.Value;

            M_Role.Status = 1;//角色状态

            //记入日志
            TOOL.Login.OptObject = M_Role.Crateator;//操作ID
            TOOL.Login.BZ += "操作后:" + M_Role.Statistics();//备注

            return M_Role;
        }





        /// <summary>
        /// 增加
        /// </summary>
        protected void Add()
        {
            try
            {


                //角色表
                BLL.B_Role B_Role_Add = new BLL.B_Role();

                Model.M_Role M_Role_Add = new Model.M_Role();

                M_Role_Add.CreateDate=System.DateTime.Now;//新增时间


                int i_Role = B_Role_Add.Add(All_M_Role(M_Role_Add));



                if (i_Role > 0)
                {

                    OperationLog OperationLogAddupdate = new OperationLog();
                    OperationLogAddupdate.OperationLogAdd("06", "03");
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>window.opener.location.reload();alert('保存成功！');window.close(); </script>");

                    //Response.Redirect("../Top.aspx");
                }
                else
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('保存失败！');</script>");
                }
            }
            catch (Exception ex)
            {
                string a = ex.Message;
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('" + a + "');</script>");

            }

        }







        /// <summary>
        /// 更新
        /// </summary>
        protected void Update(string ID)
        {
            try
            {
                //角色表
                BLL.B_Role B_Role_Update = new BLL.B_Role();

                Model.M_Role M_Role_Update = new Model.M_Role();

                M_Role_Update.CreateDate = System.DateTime.Now;//更新时间

                M_Role_Update.ID = int.Parse(ID);
                int i_Role = B_Role_Update.Update(All_M_Role(M_Role_Update));


                if (i_Role > 0)
                {

                    //记入日志
                    TOOL.Login.OptObject = M_Role_Update.ID.ToString();//操作ID
                    OperationLog OperationLogAddupdate = new OperationLog();
                    OperationLogAddupdate.OperationLogAdd("06", "04");
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>window.opener.location.reload();alert('更新成功！');window.close(); </script>");
             
                }
                else
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('更新失败！');</script>");
                }
            }

            catch (Exception ex)
            {
                string a = ex.Message;
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('" + a + "');</script>");

            }



        }





        //保存
        protected void Button2_Click1(object sender, EventArgs e)
        {

           string s_Power= this.HiddenPower.Value;


           string ID = Request.QueryString["ID"];

           try
           {
               if (string.IsNullOrEmpty(ID))
               {

                   Add();

               }

               else
               {
                   Update(ID);

               }
           }

           catch (Exception ex)
           {
               string a = ex.Message;

               Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('" + a + "');</script>");

           }


        }



    }
}