using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong.Equipment
{
    public partial class Equipment_Register_Add : System.Web.UI.Page
    {
        string ID;
        string s_BMDM = null;
        string s_SJBM = null;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
             ID = Request.QueryString["ID"];

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



                //AllocateState设备状态
                BLL.B_Dev_AllocateState dr_Dev_AllocateState = new BLL.B_Dev_AllocateState();
                this.Dr_AllocateState.DataSource = dr_Dev_AllocateState.GetList();
                this.Dr_AllocateState.DataTextField = "StateName";
                this.Dr_AllocateState.DataValueField = "ID";
                this.Dr_AllocateState.DataBind();
                this.Dr_AllocateState.SelectedIndex = 0;




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



                if (!string.IsNullOrEmpty(ID))
                {
                 
                    load(long.Parse(ID));

                    this.T_DevId.Enabled = false;
                    if (this.T_JYBH.Text.Trim() != "")
                    {
                        this.Button3.Visible = true;
                    }
                }


            }

        }




        protected void Page_PreRender(object sender, EventArgs e)
        {

            int i_DevId = int.Parse(this.Dr_DeviceNmae.SelectedValue);
            if (i_DevId == 8 || i_DevId == 7)
            {
                //this.T_BCJYSJ.Visible = true;
                //this.T_XCJYSJ.Visible = true;

                this.checktime.Visible = true;

            }

            else
            {
                //this.T_BCJYSJ.Visible = false;
                //this.T_XCJYSJ.Visible = false;
                this.checktime.Visible = false;
            }



            //设备配发状态
            if (this.Dr_AllocateState.SelectedValue == "1")
            {
                this.T_JYBH.Enabled = false;
                this.T_JYBH.Text = "";

                this.T_SJ.Enabled = false;
                this.T_SJ.Text = "";

                this.T_XM.Enabled = false;
                this.T_XM.Text = "";

                this.dr_first.Enabled = false;
                this.dr_second.Enabled = false;
                this.dr_three.Enabled = false;

                this.dr_second.SelectedIndex = 0;

                this.dr_three.SelectedIndex = 0;
            }
            else
            {
                this.T_JYBH.Enabled = true;
                this.T_SJ.Enabled = true;
                this.T_XM.Enabled = true;
                this.dr_first.Enabled = true;
                this.dr_second.Enabled = true;
                this.dr_three.Enabled = true;
            
            }

        }





        /// <summary>
        /// 加载
        /// </summary>
        protected void load(long ID)
        {
            BLL.Device B_Device_load = new BLL.Device();

            Model.Device M_Device_load = new Model.Device();
           M_Device_load= B_Device_load.GetModel(ID);

           this.Dr_DeviceNmae.SelectedValue = M_Device_load.DevType.ToString();

           this.T_Manufacturer.Text = M_Device_load.Manufacturer;
           this.T_DevId.Text = M_Device_load.DevId;

           //this.Dr_SBXH.SelectedItem.Text = M_Device_load.SBXH;//
           this.Dr_SBXH.Items.Insert(0, new ListItem(M_Device_load.SBXH));
           this.Dr_SBXH.SelectedValue = "";

            this.T_SBGG.Text=M_Device_load.SBGG;
            this.T_ProjName.Text=M_Device_load.ProjName;
             this.T_ProjNum.Text=  M_Device_load.ProjNum ;
            this.T_Price.Text=  M_Device_load.Price;
            this.T_XMFZR.Text=  M_Device_load.XMFZR;
            this.T_XMFZRDH.Text=  M_Device_load.XMFZRDH;
            this.T_CGSJ.Text=    M_Device_load.CGSJ.ToString();
            this.T_BXQ.Text=  M_Device_load.BXQ.ToString();
            this.T_BFQX.Text = M_Device_load.BFQX.ToString();
            this.Dr_State.SelectedValue = M_Device_load.WorkState.ToString();
            this.T_BCJYSJ.Text = M_Device_load.BCJYSJ.ToString();
            this.T_XCJYSJ.Text = M_Device_load.XCJYSJ.ToString();

            this.Dr_AllocateState.SelectedValue = M_Device_load.AllocateState.ToString();


            BLL.Entity B_Entity_Load = new BLL.Entity();

            this.dr_second.SelectedValue = (string)B_Entity_Load.ExistsSJBM(M_Device_load.BMDM);
            if (!string.IsNullOrEmpty(this.dr_second.SelectedItem.Text) && this.dr_second.SelectedItem.Text != "-请选择-")
            {
                dr_second_TextChanged();
               
                this.dr_three.SelectedValue = M_Device_load.BMDM; //三级

            }

            else
            {
                this.dr_first.SelectedValue = (string)B_Entity_Load.ExistsSJBM(M_Device_load.BMDM);
                if (!string.IsNullOrEmpty(this.dr_first.SelectedItem.Text))
                {
                    this.dr_second.SelectedValue = M_Device_load.BMDM;//二级


                }
                else
                {
                    this.dr_first.SelectedValue = M_Device_load.BMDM;//一级
                
                }

            
            }
        


            //ACL_USER 表
            BLL.B_ACL_USER B_ACL_USER_load = new BLL.B_ACL_USER();
            Model.M_ACL_USER M_ACL_USER_load = new Model.M_ACL_USER();

           M_ACL_USER_load= B_ACL_USER_load.GetLittleModel(M_Device_load.JYBH);


           if (string.IsNullOrEmpty(M_ACL_USER_load.JYBH))//当警员表里没有JYBH但是设备表里有JYBH
           {
               BLL.Device B_Device_Load = new BLL.Device();

               if (B_Device_Load.ExistsJYBH(this.T_JYBH.Text.Trim()))
               {

                   return;
               }
               else
               {
                   this.T_XM.Text = "";
                   this.T_JYBH.Text = "";
                   this.T_SJ.Text = "";
                   this.dr_second.SelectedIndex = 0;
                   this.dr_three.SelectedIndex = 0;
                   return;
               }
           }


           this.T_XM.Text = M_ACL_USER_load.XM;
           this.T_JYBH.Text = M_ACL_USER_load.JYBH;
           this.T_SJ.Text = M_ACL_USER_load.SJ;

           //ViewState["ID"] = M_ACL_USER_load.ID;


           //记入日志
           //TOOL.Login.OptObject = this.T_DevId.Text.Trim();//设备编号
           TOOL.Login.BZ = "操作前:" + M_Device_load.Statistics();//备注



        }

  






        /// <summary>
        /// 更新
        /// </summary>
        protected void Update( string ID)
        {
            try
            {
                //Device表
                BLL.Device B_Device_Update = new BLL.Device();
                Model.Device M_Device_Update = new Model.Device();

                M_Device_Update.ID = long.Parse(ID);




                //if (B_Device_Update.Exists(this.T_DevId.Text.Trim()))//设备编号重复
                //{
                //    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('设备编号重复');</script>");
                //    return;
                //}

                int i_Device = B_Device_Update.Update(All_Device(M_Device_Update));


                if (i_Device > 0)
                {

                    OperationLog OperationLogAddupdate = new OperationLog();
                    OperationLogAddupdate.OperationLogAdd("01", "04");
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

        /// <summary>
        /// 增加
        /// </summary>
        protected void Add()
        {
            try
            {
                //Device表
                BLL.Device Add_B_Device = new BLL.Device();

                Model.Device Add_M_Device = new Model.Device();

                TOOL.Login.BZ = "";
                //Add_M_Device.AllocateState = 2;//未回收




                if (Add_B_Device.Exists(this.T_DevId.Text.Trim()))//设备编号重复
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('设备编号重复！'); </script>");

                    return;

                }


                int i_Device = Add_B_Device.Add(All_Device(Add_M_Device));



                if (i_Device > 0)
                {



                    OperationLog OperationLogAddupdate = new OperationLog();
                    OperationLogAddupdate.OperationLogAdd("01", "03");

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

        //保存
        protected void Button2_Click(object sender, EventArgs e)
        {

           string ID=  Request.QueryString["ID"];

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
               string a= ex.Message;

               Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('" + a + "');</script>");

            }


        }


        /// <summary>
        /// 解除绑定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Button3_Click(object sender, EventArgs e)
        {
            try
            {

                //Device表
                BLL.Device B_Device_Update = new BLL.Device();
                Model.Device M_Device_Update = new Model.Device();

                M_Device_Update.ID = long.Parse(Request.QueryString["ID"]);

                M_Device_Update.JYBH = "";
                int i_Device = B_Device_Update.Update_JYBH(M_Device_Update);

                if (i_Device > 0)
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>window.location.reload();alert('解绑成功!');window.close();</script>");
                    this.T_JYBH.Text = "";
                    this.T_SJ.Text = "";
                    this.T_XM.Text = "";
                }

            }

            catch (Exception ex)
            {
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('"+ex.Message+"');</script>");
            }

        }



        /// <summary>
        /// 模型公用 Device 设备表
        /// </summary>
        /// <param name="M_Device"></param>
        /// <returns></returns>
        protected Model.Device All_Device(Model.Device M_Device)
        {

         
                M_Device.DevType = int.Parse(this.Dr_DeviceNmae.SelectedValue);

                M_Device.Manufacturer = this.T_Manufacturer.Text.Trim();


                M_Device.DevId = this.T_DevId.Text.Trim();
                M_Device.SBXH = this.Dr_SBXH.SelectedItem.Text;
                M_Device.SBGG = this.T_SBGG.Text.Trim();
                M_Device.ProjName = this.T_ProjName.Text.Trim();
                M_Device.ProjNum = this.T_ProjNum.Text.Trim();
                M_Device.Price = this.T_Price.Text.Trim();
                M_Device.XMFZR = this.T_XMFZR.Text.Trim();
                M_Device.XMFZRDH = this.T_XMFZRDH.Text.Trim();
                M_Device.CGSJ = string.IsNullOrEmpty(this.T_CGSJ.Text.Trim()) == true ? Convert.ToDateTime(null) : Convert.ToDateTime(this.T_CGSJ.Text.Trim());
                M_Device.BXQ = string.IsNullOrEmpty(this.T_BXQ.Text.Trim()) == true ? Convert.ToDateTime(null) : Convert.ToDateTime(this.T_BXQ.Text.Trim());
                M_Device.BFQX = string.IsNullOrEmpty(this.T_BFQX.Text.Trim()) == true ? Convert.ToDateTime(null) : Convert.ToDateTime(this.T_BFQX.Text.Trim());
                M_Device.WorkState = int.Parse(this.Dr_State.SelectedValue);

                M_Device.AllocateState = int.Parse(this.Dr_AllocateState.SelectedValue);
                if (string.IsNullOrEmpty(this.T_BCJYSJ.Text.Trim()))
                {
                    M_Device.BCJYSJ = null;
                }
                else
                {
                    M_Device.BCJYSJ = Convert.ToDateTime(this.T_BCJYSJ.Text.Trim());//转换不了“”值
                }
                if (string.IsNullOrEmpty(this.T_XCJYSJ.Text.Trim()))
                {
                    M_Device.XCJYSJ = null;
                }
                else
                {
                    M_Device.XCJYSJ = Convert.ToDateTime(this.T_XCJYSJ.Text.Trim());
                }
                M_Device.JYBH = this.T_JYBH.Text.Trim();



                M_Device.BMDM = this.dr_three.SelectedValue;//没选是空的
                M_Device.CreatDate = System.DateTime.Now;

                 //记入日志
                TOOL.Login.OptObject = M_Device.DevId;//设备编号
                TOOL.Login.BZ +="操作后:"+ M_Device.Statistics();//备注



                return M_Device;




        }

        /// <summary>
        /// 模型公用 ACL_USER 单位表
        /// </summary>
        /// <param name="M_ACL_USER"></param>
        /// <returns></returns>
        //protected Model.M_ACL_USER All_ACL_USER(Model.M_ACL_USER M_ACL_USER)
        //{
        //    M_ACL_USER.XM = this.T_XM.Text.Trim();
        //    M_ACL_USER.BMDM = this.dr_three.SelectedValue;
        //    M_ACL_USER.JYBH = this.T_JYBH.Text.Trim();
        //    M_ACL_USER.SJ = this.T_SJ.Text.Trim();
        //    return M_ACL_USER;
        
        
        //}

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
        /// 项目名称
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void T_ProjName_TextChanged(object sender, EventArgs e)
        {
            try
            {

                BLL.B_ProjectInfo B_Load_ProjectInfo = new BLL.B_ProjectInfo();

                Model.M_ProjectInfo M_Load_ProjectInfo = B_Load_ProjectInfo.GetModel(this.T_ProjName.Text.Trim());

                this.T_ProjName.Text = M_Load_ProjectInfo.XMMC;
                this.T_ProjNum.Text = M_Load_ProjectInfo.XMBH;
                this.T_XMFZR.Text = M_Load_ProjectInfo.XMFZR;
                this.T_XMFZRDH.Text = M_Load_ProjectInfo.XMFZRDH;
                this.T_CGSJ.Text = M_Load_ProjectInfo.CGSJ.ToString();
                this.T_BXQ.Text = M_Load_ProjectInfo.BXQ.ToString();
                this.T_Manufacturer.Text = M_Load_ProjectInfo.Manufacturer;
                this.T_BCJYSJ.Text = M_Load_ProjectInfo.BCJYSJ.ToString();
                this.T_XCJYSJ.Text = M_Load_ProjectInfo.XCJYSJ.ToString();

                //加载设备型号
                BLL.B_ProjectDetail B_Load_ProjectDetail = new BLL.B_ProjectDetail();
                DataTable D_ProjectDetail = B_Load_ProjectDetail.GetList(M_Load_ProjectInfo.XMBH);

                this.Dr_SBXH.DataSource = D_ProjectDetail;
                this.Dr_SBXH.DataTextField = "SBXH";
                this.Dr_SBXH.DataValueField = "ID";
                this.Dr_SBXH.DataBind();
                this.Dr_SBXH.SelectedIndex = 0;

                //BLL.B_ProjectDetail B_DR_ProjectDetail = new BLL.B_ProjectDetail();
                Model.M_ProjectDetail M_DR_ProjectDetail = B_Load_ProjectDetail.GetModel(int.Parse(this.Dr_SBXH.SelectedValue));
                this.T_SBGG.Text = M_DR_ProjectDetail.SBGG;
                this.T_Price.Text = M_DR_ProjectDetail.Price.ToString();
            }

            catch (Exception ex)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('项目名称不存在请重新填写!');</script>");
            }

        }

        protected void Dr_SBXH_SelectedIndexChanged(object sender, EventArgs e)
        {

           BLL.B_ProjectDetail B_DR_ProjectDetail = new BLL.B_ProjectDetail();
           Model.M_ProjectDetail M_DR_ProjectDetail= B_DR_ProjectDetail.GetModel(int.Parse(this.Dr_SBXH.SelectedValue));
           this.T_SBGG.Text = M_DR_ProjectDetail.SBGG;
           this.T_Price.Text = M_DR_ProjectDetail.Price.ToString();

        }


        //使用人警号
        protected void T_JYBH_TextChanged(object sender, EventArgs e)
        {
            try
            {

                //ACL_USER 表
                BLL.B_ACL_USER B_ACL_USER_load = new BLL.B_ACL_USER();
                Model.M_ACL_USER M_ACL_USER_load = new Model.M_ACL_USER();

                M_ACL_USER_load = B_ACL_USER_load.GetLittleModel(this.T_JYBH.Text.Trim());//先去警员表里查找警员信息

                if (string.IsNullOrEmpty(M_ACL_USER_load.JYBH))//当警员表里没有JYBH但是设备表里有JYBH
                {
                    BLL.Device B_Device_Load = new BLL.Device();

                    if (B_Device_Load.ExistsJYBH(this.T_JYBH.Text.Trim()))
                    {


                        return;
                    }

          

                }

                this.T_XM.Text = M_ACL_USER_load.XM;
                this.T_JYBH.Text = M_ACL_USER_load.JYBH;
                this.T_SJ.Text = M_ACL_USER_load.SJ;


                BLL.Entity B_Entity_Load = new BLL.Entity();

                this.dr_second.SelectedValue = (string)B_Entity_Load.ExistsSJBM(M_ACL_USER_load.BMDM);
                if (!string.IsNullOrEmpty(this.dr_second.SelectedItem.Text))
                {
                    dr_second_TextChanged();
                    this.dr_three.SelectedValue = M_ACL_USER_load.BMDM; //三级

                }

                else
                {
                    this.dr_first.SelectedValue = (string)B_Entity_Load.ExistsSJBM(M_ACL_USER_load.BMDM);
                    if (!string.IsNullOrEmpty(this.dr_first.SelectedItem.Text))
                    {
                        this.dr_second.SelectedValue = M_ACL_USER_load.BMDM;//二级


                    }
                    else
                    {
                        this.dr_first.SelectedValue = M_ACL_USER_load.BMDM;//一级

                    }


                }
        


            }
            catch (Exception ex)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert(" + ex.Message+ ");</script>");
            }

        }

        //设备名称  DevId 设备编号  只有酒精测试仪和测速仪才有检验时间
        protected void Dr_DeviceNmae_SelectedIndexChanged(object sender, EventArgs e)
        {
            int i_DevId = int.Parse(this.Dr_DeviceNmae.SelectedValue);

            if (i_DevId == 8 || i_DevId == 7)
            {
                //this.T_BCJYSJ.Visible = true;
                //this.T_XCJYSJ.Visible = true;

                this.checktime.Visible = true;

            }

            else
            {
                //this.T_BCJYSJ.Visible = false;
                //this.T_XCJYSJ.Visible = false;
                this.checktime.Visible = false;
            }


        }




    }
}