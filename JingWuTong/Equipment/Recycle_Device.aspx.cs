using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong.Recycle
{
    public partial class Recycle_Device : System.Web.UI.Page
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

                if (s_power[12] == "0") //删除
                {
                    this.GridView1.Columns[11].Visible = false;


                }

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


                //设备状态
                BLL.B_Dev_WorkState dr_Dev_WorkState = new BLL.B_Dev_WorkState();
                this.Dr_State.DataSource = dr_Dev_WorkState.GetList();
                this.Dr_State.DataTextField = "StateName";
                this.Dr_State.DataValueField = "ID";
                this.Dr_State.DataBind();
                //this.Dr_State.Items.Insert(0, new ListItem("---请选择---", "-1"));
                this.Dr_State.Items.Insert(0, new ListItem("全部", "-1"));
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

            DataTable ds_Source = F_Entity.GetList(this.dr_second.SelectedValue,"");

            this.dr_three.DataSource = F_Entity.GetList(this.dr_second.SelectedValue,"");
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



        //删除
        protected void GridView1_RowCommand(object sender, GridViewCommandEventArgs e)
        {

            switch (e.CommandName)
            {
                case "MyDel":
                    BLL.B_DeviceLog B_DeviceLog_Del = new BLL.B_DeviceLog();



                    string[] estr = e.CommandArgument.ToString().Split(',');

                    long l_ParentID = long.Parse(estr[0]);
                    string i_DevId = estr[1];

                    //删除DeviceLog表中数据
                    bool b_ID = B_DeviceLog_Del.Delete(l_ParentID);

                    //更新Device表中数据
                    BLL.Device B_Device_Update = new BLL.Device();

                  int I_WDevId = B_Device_Update.UpdateAllocateState_WDevId(i_DevId);

                  if (b_ID == true && I_WDevId>0)
                    {

                        OperationLog OperationLogAddupdate = new OperationLog();
                        OperationLogAddupdate.OperationLogAdd("03", "04");

                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('删除成功！');</script>");
                        this.GridView1.DataBind();

                    }
                    else
                    {
                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('删除失败！');</script>");

                    }
                    break;

                default:

                    break;
            }


        }




        protected void Button2_Click(object sender, EventArgs e)
        {

            this.Dr_State.Items[0].Value = "-2";
            GridView1.PageIndex = 0;

        }



        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnOutExcel_Click(object sender, EventArgs e)
        {

           // int i_DeviceNmae = int.Parse(this.Dr_DeviceNmae.SelectedValue);//设备名称

            string s_strat = this.T_strat.Text.Trim();//采购开始时间
            string s_now = this.T_now.Text.Trim();


            int i_State = int.Parse(this.Dr_State.SelectedValue);//设备状态

            string s_three = this.dr_three.SelectedValue;//所属部门

            string s_search= this.T_search.Text.Trim();

            string s_second = this.dr_second.SelectedValue;

            DataTable thisTable = null;


            BLL.B_DeviceLog B_DeviceLog_Out = new BLL.B_DeviceLog();

            thisTable = B_DeviceLog_Out.OutExcelRecycle(s_three,s_second, s_strat, s_now,i_State, s_search);

            if (thisTable.Rows.Count <= 0)
            {
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('没有数据！！！');</script>");
                return;

            }



            if (thisTable != null)
            {
                StringWriter sw = new StringWriter();
                //生成列
                sw.WriteLine("日志类型\t所属部门\t联系人\t联系电话\t警号\t回收名称\t设备编号\t回收日期\t备注");
                foreach (DataRow dr in thisTable.Rows)
                {

                    //生成行
                    sw.WriteLine(dr["LogType"] + "\t" + dr["BMJC"] + "\t" + dr["LXR"] + "\t" + dr["LXDH"] + "\t" + dr["JYBH"] + "\t" + dr["TypeName"] + "\t" + dr["DevId"] + "\t" + dr["LogTime"] + "\t" + dr["BZ"]);
                }
                sw.Close();

                //记入日志
                TOOL.Login.OptObject = "";//操作ID
                TOOL.Login.BZ = "导出了" + thisTable.Rows.Count + "条数据";//备注
                OperationLog OperationLogAddupdate = new OperationLog();
                OperationLogAddupdate.OperationLogAdd("03", "11");


                Response.Clear();
                string fileName = "回收统计(" + DateTime.Now.ToString("yyyy-MM-dd") + ").xls";

                Response.AddHeader("Content-Disposition", "attachment;  filename=" + HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8));

                Response.ContentType = "application/ms-excel";
                Response.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");
                Response.Write(sw);
                Response.End();
            }





        }



    }
}