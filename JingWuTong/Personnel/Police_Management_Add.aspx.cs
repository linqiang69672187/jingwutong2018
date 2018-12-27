using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong.Personnel
{
    public partial class Police_Management_Add : System.Web.UI.Page
    {
       
        protected void Page_Load(object sender, EventArgs e)
        {


            if (!Page.IsPostBack)
            {

                string ID= Request.QueryString["ID"];

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

                //警员类型
                BLL.B_PoliceType dr_PoliceType = new BLL.B_PoliceType();

                this.Dr_JYLX.DataSource = dr_PoliceType.GetList();
                this.Dr_JYLX.DataTextField = "TypeName";
                this.Dr_JYLX.DataValueField = "ID";
                this.Dr_JYLX.DataBind();

                //this.Dr_JYLX.Items.Insert(0, new ListItem("---选择类型---", "-1"));
                this.Dr_JYLX.SelectedIndex = 0;

                //角色类型

                BLL.B_Role dr_Role = new BLL.B_Role();
                this.Dr_JSID.DataSource = dr_Role.GetList();
                this.Dr_JSID.DataTextField = "RoleName";
                this.Dr_JSID.DataValueField = "ID";
                this.Dr_JSID.DataBind();

                //this.Dr_JSID.Items.Insert(0, new ListItem("---选择类型---", "-1"));
                this.Dr_JSID.SelectedIndex = 0;


                //警员职务

                BLL.Position dr_Position = new BLL.Position();

                this.Dr_LDJB.DataSource = dr_Position.GetList();
                this.Dr_LDJB.DataTextField = "PositionName";
                this.Dr_LDJB.DataValueField = "ID";
                this.Dr_LDJB.DataBind();
                //this.Dr_JSID.Items.Insert(0, new ListItem("---选择类型---", "-1"));
                this.Dr_LDJB.SelectedIndex = 0;


                if (!string.IsNullOrEmpty(ID))
                {
                    this.L_head.InnerText = "修改警员";
                    this.T_JYBH.Enabled = false;
                    load(long.Parse(ID));
                }
                else
                {
                    this.L_head.InnerText = "新增警员";
                }


            }

        }




        /// <summary>
        /// 加载
        /// </summary>
        protected void load(long ID)
        {

            BLL.B_ACL_USER B_ACL_USER_Load = new BLL.B_ACL_USER();

            Model.M_ACL_USER M_ACL_USER_Load = new Model.M_ACL_USER();

           M_ACL_USER_Load= B_ACL_USER_Load.GetModel(ID);

           this.T_XM.Text=M_ACL_USER_Load.XM;
          this.T_SJ.Text=M_ACL_USER_Load.SJ;
          this.T_JYBH.Text= M_ACL_USER_Load.JYBH ;
          this.T_SFZMHM.Text= M_ACL_USER_Load.SFZMHM ;
          this.Dr_JSID.SelectedValue= M_ACL_USER_Load.JSID.ToString() ;
          this.Dr_JYLX.SelectedValue=M_ACL_USER_Load.JYLX.ToString() ;
          this.Dr_LDJB.SelectedValue = M_ACL_USER_Load.LDJB.ToString();
       




          BLL.Entity B_Entity_Load = new BLL.Entity();

          this.dr_second.SelectedValue = (string)B_Entity_Load.ExistsSJBM(M_ACL_USER_Load.BMDM);
          if (!string.IsNullOrEmpty(this.dr_second.SelectedItem.Text) && this.dr_second.SelectedItem.Text != "-请选择-")
          { 
              dr_second_TextChanged();

              this.dr_three.SelectedValue = M_ACL_USER_Load.BMDM; //三级

          }

          else
          {
              this.dr_first.SelectedValue = (string)B_Entity_Load.ExistsSJBM(M_ACL_USER_Load.BMDM);
              if (!string.IsNullOrEmpty(this.dr_first.SelectedItem.Text) )
              {
                  this.dr_second.SelectedValue = M_ACL_USER_Load.BMDM;//二级


              }
              else
              {
                  this.dr_first.SelectedValue = M_ACL_USER_Load.BMDM;//一级

              }


          }


          //记入日志
          //TOOL.Login.OptObject = this.T_DevId.Text.Trim();//设备编号
          TOOL.Login.BZ = "操作前:" + M_ACL_USER_Load.Statistics();//备注

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
        /// 模型公用 ACL_USER 警员表
        /// </summary>
        /// <param name="M_DeviceLog"></param>
        /// <returns></returns>
        protected Model.M_ACL_USER All_M_ACL_USER(Model.M_ACL_USER M_ACL_USER)
        {

            M_ACL_USER.XM = this.T_XM.Text.Trim();
            M_ACL_USER.SJ = this.T_SJ.Text.Trim();

            M_ACL_USER.JYBH = this.T_JYBH.Text.Trim();

            M_ACL_USER.SFZMHM = this.T_SFZMHM.Text.Trim();
            M_ACL_USER.JSID = int.Parse(this.Dr_JSID.SelectedValue);
            M_ACL_USER.JYLX = int.Parse(this.Dr_JYLX.SelectedValue);
            M_ACL_USER.BMDM = this.dr_three.SelectedValue;
            M_ACL_USER.LDJB = this.Dr_LDJB.SelectedValue;


            //M_ACL_USER.CJSJ= System.DateTime.Now;

            //记入日志
            TOOL.Login.OptObject = M_ACL_USER.JYBH;//操作ID
            TOOL.Login.BZ += "操作后:" + M_ACL_USER.Statistics();//备注


            return M_ACL_USER;
        }





        /// <summary>
        /// 增加
        /// </summary>
        protected void Add()
        {
            try
            {
                //ACL_USER 警员表
                BLL.B_ACL_USER B_ACL_USER_Add = new BLL.B_ACL_USER();

                Model.M_ACL_USER M_ACL_USER_Add = new Model.M_ACL_USER();
                TOOL.Login.BZ = "";

                if (B_ACL_USER_Add.ExistsJYBH(this.T_JYBH.Text.Trim())) //不能重复
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('警员编号重复！');</script>");

                    return;
                }

                M_ACL_USER_Add.CJSJ = System.DateTime.Now;//新增时间


                int i_ACL_USER = B_ACL_USER_Add.AddACL_USER(All_M_ACL_USER(M_ACL_USER_Add));

                if (i_ACL_USER > 0)
                {
                    OperationLog OperationLogAddupdate = new OperationLog();
                    OperationLogAddupdate.OperationLogAdd("05", "03");
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
                  //ACL_USER 警员表
                BLL.B_ACL_USER B_ACL_USER_Update = new BLL.B_ACL_USER();

                Model.M_ACL_USER M_ACL_USER_Update = new Model.M_ACL_USER();

                M_ACL_USER_Update.ID =long.Parse(ID);

                M_ACL_USER_Update.GXSJ = System.DateTime.Now;//更新时间

                int i_ACL_USER = B_ACL_USER_Update.UpdateACL_USE(All_M_ACL_USER(M_ACL_USER_Update));


                if (i_ACL_USER > 0)
                {
                    OperationLog OperationLogAddupdate = new OperationLog();
                    OperationLogAddupdate.OperationLogAdd("05", "04");
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>window.opener.location.reload();alert('更新成功！');window.close();</script>");

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