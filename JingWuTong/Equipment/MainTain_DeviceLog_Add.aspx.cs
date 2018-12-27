using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong.Maintain
{
    public partial class MainTain_DeviceLog_Add : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {


                //设备名称
                BLL.B_DeviceType dr_DeviceType = new BLL.B_DeviceType();
                this.Dr_DeviceNmae.DataSource = dr_DeviceType.GetList();
                this.Dr_DeviceNmae.DataTextField = "TypeName";
                this.Dr_DeviceNmae.DataValueField = "ID";
                this.Dr_DeviceNmae.DataBind();
                this.Dr_DeviceNmae.SelectedIndex = 0;



                //设备状态
                BLL.B_Dev_WorkState dr_Dev_WorkState = new BLL.B_Dev_WorkState();
                this.Dr_State.DataSource = dr_Dev_WorkState.GetList();
                this.Dr_State.DataTextField = "StateName";
                this.Dr_State.DataValueField = "ID";
                this.Dr_State.DataBind();
                this.Dr_State.SelectedIndex = 0;


                string s_BMDM = null;
                string s_SJBM = null;

                if (HttpContext.Current.Request.Cookies["cookieName"] != null)
                {

                    HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];

                    if (cookies["BMDM"] != null)
                    {
                        s_BMDM = Server.UrlDecode(cookies["BMDM"]);
                        s_SJBM = Server.UrlDecode(cookies["SJBM"]);
                    }


                }



                //使用人单位
                BLL.Entity F_Entity = new BLL.Entity();



                this.dr_first.DataSource = F_Entity.GetList("330000000000", "");
                this.dr_first.DataTextField = "BMMC";
                this.dr_first.DataValueField = "BMDM";
                this.dr_first.DataBind();
                this.dr_first.SelectedIndex = 0;


                if (s_SJBM == "330000000000")
                {
                    this.dr_second.DataSource = F_Entity.GetList(this.dr_first.SelectedValue, "");
                    this.dr_second.DataTextField = "BMMC";
                    this.dr_second.DataValueField = "BMDM";
                    this.dr_second.DataBind();
                    this.dr_second.Items.Insert(0, new ListItem("-请选择-", ""));
                    this.dr_second.SelectedIndex = 0;

                    this.dr_three.DataSource = F_Entity.GetList(this.dr_second.SelectedValue, "");
                    this.dr_three.DataTextField = "BMMC";
                    this.dr_three.DataValueField = "BMDM";
                    this.dr_three.DataBind();
                    this.dr_three.Items.Insert(0, new ListItem("-请选择-", ""));
                    this.dr_three.SelectedIndex = 0;

                }

                else
                {

                    if (s_SJBM != this.dr_first.SelectedValue)
                    {

                        this.dr_second.DataSource = F_Entity.GetList(this.dr_first.SelectedValue, s_SJBM);
                        this.dr_second.DataTextField = "BMMC";
                        this.dr_second.DataValueField = "BMDM";
                        this.dr_second.DataBind();
                        this.dr_second.SelectedIndex = 0;

                        this.dr_three.DataSource = F_Entity.GetList(this.dr_second.SelectedValue, s_BMDM);
                        this.dr_three.DataTextField = "BMMC";
                        this.dr_three.DataValueField = "BMDM";
                        this.dr_three.DataBind();
                        this.dr_three.SelectedIndex = 0;
                    }
                    else
                    {
                        this.dr_second.DataSource = F_Entity.GetList(this.dr_first.SelectedValue, s_BMDM);
                        this.dr_second.DataTextField = "BMMC";
                        this.dr_second.DataValueField = "BMDM";
                        this.dr_second.DataBind();
                        this.dr_second.SelectedIndex = 0;



                        this.dr_three.DataSource = F_Entity.GetList(this.dr_second.SelectedValue, "");
                        this.dr_three.DataTextField = "BMMC";
                        this.dr_three.DataValueField = "BMDM";
                        this.dr_three.DataBind();
                        this.dr_three.Items.Insert(0, new ListItem("-请选择-", ""));
                        this.dr_three.SelectedIndex = 0;

                    }

                }


                if (!string.IsNullOrEmpty(Request.QueryString["ID"]))
                {

                    long ID = long.Parse(Request.QueryString["ID"]);
                    load(ID);

                }

                       

            }
        }


        //使用人警号
        protected void T_JYBH_TextChanged(object sender, EventArgs e)
        {

            //ACL_USER 表
            BLL.B_ACL_USER B_ACL_USER_load = new BLL.B_ACL_USER();
            Model.M_ACL_USER M_ACL_USER_load = new Model.M_ACL_USER();

            M_ACL_USER_load = B_ACL_USER_load.GetLittleModel(this.T_JYBH.Text.Trim());

            this.T_XM.Text = M_ACL_USER_load.XM;
            this.T_JYBH.Text = M_ACL_USER_load.JYBH;
            this.T_SJ.Text = M_ACL_USER_load.SJ;

        }



        //使用人单位
        protected void dr_second_SelectedIndexChanged(object sender, EventArgs e)
        {

            BLL.Entity F_Entity = new BLL.Entity();


            DataTable ds_Source = F_Entity.GetList(this.dr_second.SelectedValue,"");

            this.dr_three.DataSource = F_Entity.GetList(this.dr_second.SelectedValue,"");
            this.dr_three.DataTextField = "BMMC";
            this.dr_three.DataValueField = "BMDM";
            this.dr_three.DataBind();

            if (ds_Source.Rows.Count > 0)
                this.dr_three.SelectedIndex = 0;


        }


        
        /// <summary>
        /// 加载
        /// </summary>
        protected void load(long ID)
        {
            BLL.B_DeviceLog B_DeviceLog_Load = new BLL.B_DeviceLog();
            Model.M_DeviceLog M_DeviceLog_Load = new Model.M_DeviceLog();

            M_DeviceLog_Load = B_DeviceLog_Load.GetModel(ID);


            this.Dr_DeviceNmae.SelectedValue = M_DeviceLog_Load.DevType.ToString();
            this.T_DevId.Text = M_DeviceLog_Load.DevId;
            this.Dr_State.SelectedValue = M_DeviceLog_Load.DevState.ToString();

            this.T_JYBH.Text = M_DeviceLog_Load.JYBH;
            this.T_XM.Text = M_DeviceLog_Load.BXR;
            this.T_SJ.Text = M_DeviceLog_Load.Tel;
            this.T_BZ.Text = M_DeviceLog_Load.BZ;


            BLL.Entity B_Entity_Load = new BLL.Entity();

            this.dr_second.SelectedValue = (string)B_Entity_Load.ExistsSJBM(M_DeviceLog_Load.Entity);
            if (!string.IsNullOrEmpty(this.dr_second.SelectedItem.Text) && this.dr_second.SelectedItem.Text != "-请选择-")
            {
             dr_second_TextChanged();

                this.dr_three.SelectedValue = M_DeviceLog_Load.Entity; //三级

            }

            else
            {
                this.dr_first.SelectedValue = (string)B_Entity_Load.ExistsSJBM(M_DeviceLog_Load.Entity);
                if (!string.IsNullOrEmpty(this.dr_first.SelectedItem.Text))
                {
                    this.dr_second.SelectedValue = M_DeviceLog_Load.Entity;//二级


                }
                else
                {
                    this.dr_first.SelectedValue = M_DeviceLog_Load.Entity;//一级

                }


            }


            //记入日志
            //TOOL.Login.OptObject = this.T_DevId.Text.Trim();//设备编号
            TOOL.Login.BZ = "操作前:" + M_DeviceLog_Load.Statistics();//备注

        
        }



        private void dr_second_TextChanged()
        {
            BLL.Entity F_Entity = new BLL.Entity();

            DataTable ds_Source = F_Entity.GetList(this.dr_second.SelectedValue, "");

            this.dr_three.DataSource = F_Entity.GetList(this.dr_second.SelectedValue, "");
            this.dr_three.DataTextField = "BMMC";
            this.dr_three.DataValueField = "BMDM";
            this.dr_three.DataBind();
            this.dr_three.Items.Insert(0, new ListItem("-请选择-", ""));
            if (ds_Source.Rows.Count > 0)
                this.dr_three.SelectedIndex = 0;

        }



        /// <summary>
        /// 模型公用 DeviceLog 设备表
        /// </summary>
        /// <param name="M_DeviceLog"></param>
        /// <returns></returns>
        protected Model.M_DeviceLog All_M_M_DeviceLog(Model.M_DeviceLog M_DeviceLog)
        {
            M_DeviceLog.DevType = int.Parse(this.Dr_DeviceNmae.SelectedValue);
            M_DeviceLog.DevId = this.T_DevId.Text.Trim();
            M_DeviceLog.DevState = int.Parse(this.Dr_State.SelectedValue);
            M_DeviceLog.JYBH=this.T_JYBH.Text.Trim();

            if (HttpContext.Current.Request.Cookies["cookieName"] != null)
            {

                HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];
                if (cookies["name"] != null)
                {
                    M_DeviceLog.BXR =Server.UrlDecode(cookies["name"]);
                }
            }
            M_DeviceLog.Tel = this.T_SJ.Text.Trim();
            M_DeviceLog.BZ = this.T_BZ.Text.Trim();
            M_DeviceLog.Entity = this.dr_three.SelectedValue;
            M_DeviceLog.LogTime = System.DateTime.Now;



            //记入日志
            TOOL.Login.OptObject = M_DeviceLog.ID.ToString();//操作ID
            TOOL.Login.BZ += "操作后:" + M_DeviceLog.Statistics();//备注

            return M_DeviceLog;
        }



        /// <summary>
        /// 增加
        /// </summary>
        protected void Add()
        {
            try
            {
                //DeviceLog表
                BLL.B_DeviceLog B_DeviceLog_Add = new BLL.B_DeviceLog();
                Model.M_DeviceLog M_DeviceLog_Add = new Model.M_DeviceLog();
                TOOL.Login.BZ = "";
                M_DeviceLog_Add.LogType = 1;//1:表示维修记录；2:表示设备日志
                M_DeviceLog_Add.AllocateState = 2;//表示配发
                M_DeviceLog_Add.LogTime = System.DateTime.Now;

               int i_DeviceLog= B_DeviceLog_Add.Add(All_M_M_DeviceLog(M_DeviceLog_Add));

               if (i_DeviceLog > 0)
                {
                    OperationLog OperationLogAddupdate = new OperationLog();
                    OperationLogAddupdate.OperationLogAdd("02", "03");

                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>window.opener.location.reload();alert('保存成功！');window.close(); </script>");

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
                //DeviceLog表
                BLL.B_DeviceLog B_DeviceLog_Update = new BLL.B_DeviceLog();
                Model.M_DeviceLog M_DeviceLog_Update = new Model.M_DeviceLog();

                M_DeviceLog_Update.LogTime = System.DateTime.Now;
                M_DeviceLog_Update.ID = long.Parse(ID);
                int i_DeviceLog = B_DeviceLog_Update.Update(All_M_M_DeviceLog(M_DeviceLog_Update));


             if (i_DeviceLog > 0)
                {

                    OperationLog OperationLogAddupdate = new OperationLog();
                    OperationLogAddupdate.OperationLogAdd("02", "04");

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
        protected void Button2_Click(object sender, EventArgs e)
        {

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