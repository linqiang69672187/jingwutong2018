using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong.Remind
{
    public partial class Remind_Device: System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {



            if (!Page.IsPostBack)
            {


                BLL.B_Role B_Role_Load = new BLL.B_Role();
                string[] s_power = ((string)B_Role_Load.Exists(TOOL.Login.I_JSID)).Split('|');//权限参数

                //if (s_power[11] == "0")
                //{


                //    this.B_Add.Visible = false;  //新增

                //}

                //if (s_power[12] == "0") //删除
                //{
                //    this.GridView1.Columns[11].Visible = false;

                //    //    System.Web.UI.HtmlControls.HtmlImage Imagedel = (System.Web.UI.HtmlControls.HtmlImage)this.GridView1.Rows[i].FindControl("img_del");
                //}

                if (s_power[13] == "0")//编辑
                {
                    this.GridView1.Columns[10].Visible = false;
                }

                if (s_power[14] == "0")//搜索
                {
                    this.Button1.Visible = false;
                }

                if (s_power[15] == "0")//导出
                {
                    this.B_Out.Visible = false;
                }

                //if (s_power[16] == "0")//导入
                //{
                //    this.B_Into.Visible = false;
                //}

                if (s_power[17] == "0")//打印
                {
                    this.B_Print.Visible = false;
                }


                //this.T_strat.Text = DateTime.Now.AddDays(-15).ToString("yyyy/MM/dd HH:mm:ss");
                //this.T_now.Text = System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");



                //设备名称
                BLL.B_DeviceType dr_DeviceType = new BLL.B_DeviceType();
                this.Dr_DeviceNmae.DataSource = dr_DeviceType.GetList();
                this.Dr_DeviceNmae.DataTextField = "TypeName";
                this.Dr_DeviceNmae.DataValueField = "ID";
                this.Dr_DeviceNmae.DataBind();
                //this.Dr_DeviceNmae.Items.Insert(0, new ListItem("---请选择---", "-1"));
                this.Dr_DeviceNmae.Items.Insert(0, new ListItem("全部", "-1"));
                this.Dr_DeviceNmae.SelectedIndex = 0;

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
                    this.dr_second.Items.Insert(0, new ListItem("全部", ""));
                    this.dr_second.SelectedIndex = 0;

                    this.dr_three.DataSource = F_Entity.GetList(this.dr_second.SelectedValue, "");
                    this.dr_three.DataTextField = "BMMC";
                    this.dr_three.DataValueField = "BMDM";
                    this.dr_three.DataBind();
                    this.dr_three.Items.Insert(0, new ListItem("全部", ""));
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
                        this.dr_three.Items.Insert(0, new ListItem("全部", ""));
                        this.dr_three.SelectedIndex = 0;

                    }

                }



             
            }

        }




        //使用人单位
        protected void dr_second_SelectedIndexChanged(object sender, EventArgs e)
        {

            BLL.Entity F_Entity = new BLL.Entity();

            DataTable ds_Source = F_Entity.GetList(this.dr_second.SelectedValue, "");

            this.dr_three.DataSource = F_Entity.GetList(this.dr_second.SelectedValue, "");
            this.dr_three.DataTextField = "BMMC";
            this.dr_three.DataValueField = "BMDM";
            this.dr_three.DataBind();
            this.dr_three.Items.Insert(0, new ListItem("全部", ""));
            if (ds_Source.Rows.Count > 0)
                this.dr_three.SelectedIndex = 0;


        }


      

        /// <summary>
        /// 数据加载后运行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                if (e.Row.RowIndex % 2 == 0)
                {
                    e.Row.BackColor = Color.FromArgb(33, 51, 102);

                }
                else
                {
                    e.Row.BackColor = Color.FromArgb(19, 32, 77);
                }
            }
     

        }



        protected void Button2_Click(object sender, EventArgs e)
        {


            this.Dr_DeviceNmae.Items[0].Value = "-2";
            this.Dr_Remind.Items[0].Value = "0";
            GridView1.PageIndex = 0;

        }


        //更新前
        protected void GridView1_RowEditing(object sender, GridViewEditEventArgs e)
        {
            GridView1.EditIndex = e.NewEditIndex;
    
        }


        //更新
        protected void GridView1_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
     
            Model.Device M_Device_update=new Model.Device();

            BLL.B_DeviceType B_DeviceType = new BLL.B_DeviceType();

           

            int i_DevType = (int)B_DeviceType.GetListID(((TextBox)(GridView1.Rows[e.RowIndex].Cells[2].FindControl("T_TypeName"))).Text.ToString().Trim());

             M_Device_update.DevType = i_DevType;


             M_Device_update.Manufacturer = ((TextBox)(GridView1.Rows[e.RowIndex].Cells[3].FindControl("T_Manufacturer"))).Text.ToString().Trim();

             M_Device_update.SBXH = ((TextBox)(GridView1.Rows[e.RowIndex].Cells[3].FindControl("T_SBXH"))).Text.ToString().Trim();

   
             M_Device_update.ProjName = ((TextBox)(GridView1.Rows[e.RowIndex].Cells[4].FindControl("T_ProjName"))).Text.ToString().Trim();

             M_Device_update.CGSJ = Convert.ToDateTime(((TextBox)(GridView1.Rows[e.RowIndex].Cells[5].FindControl("T_CGSJ"))).Text.ToString().Trim());

             M_Device_update.BFQX = Convert.ToDateTime(((TextBox)(GridView1.Rows[e.RowIndex].Cells[6].FindControl("T_BFQX"))).Text.ToString().Trim());
             M_Device_update.Price = ((TextBox)(GridView1.Rows[e.RowIndex].Cells[7].FindControl("T_Price"))).Text.ToString().Trim();

             M_Device_update.CreatDate = System.DateTime.Now;

             M_Device_update.ID =long.Parse(((Label)(GridView1.Rows[e.RowIndex].Cells[1].FindControl("L_ID"))).Text.ToString().Trim());

             //TOOL.Login.BZ = "操作前:" + M_Device_update.Statistics();//备注
        

            //BLL.Device B_Device_update = new BLL.Device();
            //B_Device_update.UpdateRemind(M_Device_update);

           
             this.ObjectDataSource1.UpdateMethod = "UpdateRemind1";

            this.ObjectDataSource1.UpdateParameters.Add("DevType",DbType.Int32, M_Device_update.DevType.ToString());
            this.ObjectDataSource1.UpdateParameters.Add("Manufacturer",DbType.String, M_Device_update.Manufacturer);
            this.ObjectDataSource1.UpdateParameters.Add("SBXH",DbType.String, M_Device_update.SBXH);
            this.ObjectDataSource1.UpdateParameters.Add("ProjName",DbType.String, M_Device_update.ProjName);
            this.ObjectDataSource1.UpdateParameters.Add("CGSJ",DbType.Date, M_Device_update.CGSJ.ToString());
            this.ObjectDataSource1.UpdateParameters.Add("BFQX",DbType.Date, M_Device_update.BFQX.ToString());
            this.ObjectDataSource1.UpdateParameters.Add("Price",DbType.String, M_Device_update.Price);
            this.ObjectDataSource1.UpdateParameters.Add("CreatDate",DbType.Date, M_Device_update.CreatDate.ToString());
            this.ObjectDataSource1.UpdateParameters.Add("ID",DbType.Int64, M_Device_update.ID.ToString());

            this.ObjectDataSource1.UpdateParameters.Add("TypeName",DbType.String, "");
            this.ObjectDataSource1.UpdateParameters.Add("Remind", DbType.String, "");
              this.ObjectDataSource1.Update();

              GridView1.EditIndex = e.RowIndex;
            //GridView1.EditIndex = -1;


              //记入日志
              TOOL.Login.OptObject = M_Device_update.ID.ToString();//操作ID
              TOOL.Login.BZ = "操作后:" + M_Device_update.Statistics();//备注

              OperationLog OperationLogAddupdate = new OperationLog();
              OperationLogAddupdate.OperationLogAdd("04", "04");
        }

        //取消
        protected void GridView1_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            GridView1.EditIndex = -1;
           
        }







        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnOutExcel_Click(object sender, EventArgs e)
        {


            int s_DeviceNmae = int.Parse(this.Dr_DeviceNmae.SelectedValue);//设备名称

            int i_Remind = int.Parse(this.Dr_Remind.SelectedValue);//设备提醒

            string s_strat = this.T_strat.Text.Trim();//采购开始时间

            string s_now = this.T_now.Text.Trim();

            string s_three = this.dr_three.SelectedValue;//所属部门

            string s_search = this.T_search.Text.Trim();

            string s_second = this.dr_second.SelectedValue;

            //int i_Index= this.GridView1.PageIndex*10;

            //int i_Count =10;


            DataTable thisTable = null;

            BLL.Device B_Device = new BLL.Device();

            thisTable = B_Device.OutExcelRemind(i_Remind, s_DeviceNmae, s_strat, s_now,s_three,s_second, s_search);

            if (thisTable.Rows.Count <= 0)
            {
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('没有数据！！！');</script>");
                return;

            }



            if (thisTable != null)
            {
                StringWriter sw = new StringWriter();
                //生成列
                sw.WriteLine("设备名称\t厂家\t型号\t项目名称\t采购时间\t到期时间\t单价\t提醒类型\t创建时间");
                foreach (DataRow dr in thisTable.Rows)
                {

                    //生成行
                    sw.WriteLine(dr["TypeName"] + "\t" + dr["Manufacturer"] + "\t" + dr["SBXH"] +"\t"+ dr["ProjName"] + "\t" + dr["CGSJ"] + "\t" + dr["BFQX"] + "\t" + dr["Price"] + "\t" + dr["Remind"] + "\t" + dr["CreatDate"]);
                }
                sw.Close();

                //记入日志
                TOOL.Login.OptObject = "";//操作ID
                TOOL.Login.BZ = "导出了" + thisTable.Rows.Count + "条数据";//备注
                OperationLog OperationLogAddupdate = new OperationLog();
                OperationLogAddupdate.OperationLogAdd("04", "11");



                Response.Clear();
                string fileName = "设备提醒(" + DateTime.Now.ToString() + ").xls";
                Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}.xls", Server.UrlEncode(fileName)));
                Response.ContentType = "application/vnd.ms-excel";
                Response.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");
                Response.Write(sw);
                Response.End();
            }




        }





        //一键公用功能
        protected void btnScrap_Click(object sender, EventArgs e)
        {

            string name = "操作后:";
            TOOL.Login.BZ = "";
            if (HttpContext.Current.Request.Cookies["cookieName"] != null)
            {
                HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];
                name = cookies["name"];

            }

            int I_sum = 0;//计算被选中的条数

            int i_Device = 0;
            int i_DeviceLog = 0;

            for (int i = 0; i < this.GridView1.Rows.Count; i++)
            {



                CheckBox box = (CheckBox)this.GridView1.Rows[i].Cells[0].FindControl("CheckBox1");
                if (box.Checked == true)
                {
                    I_sum++;
                    //更新Device表中WorkState 工作状态字段
                    BLL.Device B_Device_update = new BLL.Device();
                    Model.Device M_Device_update = new Model.Device();

                    long L_ID = long.Parse(((Label)this.GridView1.Rows[i].FindControl("L_ID")).Text);

                    //long L_ID =long.Parse( this.GridView1.Rows[i].Cells[1].Text);

                        M_Device_update.ID = L_ID;
            
                        M_Device_update.WorkState = int.Parse(this.HiddenNState.Text);
                        i_Device = B_Device_update.UpdateWorkState(M_Device_update);


                


                    //如果报修成功就往DeviceLog表中插入报修记录

                    BLL.B_DeviceLog B_DeviceLog_Add = new BLL.B_DeviceLog();

                    Model.M_DeviceLog M_DeviceLog_Add = new Model.M_DeviceLog();

                    M_DeviceLog_Add.LogType = 1;//设备维修
                    M_DeviceLog_Add.LogTime = Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd"));
                    M_DeviceLog_Add.BXR = name;
                    M_DeviceLog_Add.BZ = this.T_BZ.Text.Trim();
                    M_DeviceLog_Add.AllocateState = 2;//配发状态 未回收
                    M_DeviceLog_Add.DevState = int.Parse(this.HiddenNState.Text);//设备状态
                    i_DeviceLog = B_DeviceLog_Add.AddThree(M_DeviceLog_Add);


                    TOOL.Login.OptObject += M_Device_update.ID.ToString()+",";//操作ID
                    TOOL.Login.BZ += "AllocateState:" + M_Device_update.AllocateState.ToString()+",";//备注

                }

            }

            if (i_Device > 0 && i_DeviceLog > 0)
            {



                OperationLog OperationLogAddupdate = new OperationLog();
                OperationLogAddupdate.OperationLogAdd("04", "09");
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('操作成功');</script>");
                return;
            }



            if (I_sum <= 0)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('请先选择要操作的行！！！');</script>");
                return;
            }



        }





    }
}