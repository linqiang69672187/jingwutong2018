using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong.SystemSetup
{
    public partial class System_Log : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {

                BLL.B_Role B_Role_Load = new BLL.B_Role();
                string[] s_power = ((string)B_Role_Load.Exists(TOOL.Login.I_JSID)).Split('|');//权限参数

       
                if (s_power[28] == "0")//搜索
                {
                    this.Button1.Visible = false;
                }


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

//---------------------------------------------------------------------------

                string temp=e.Row.Cells[7].Text;
                e.Row.Cells[7].ToolTip = temp;

                e.Row.Cells[7].Text ="查看详情";

            }

           


        }


        //protected void GridView1_RowCreated(object sender, GridViewRowEventArgs e)
        //{
        //    if (e.Row.RowType == DataControlRowType.DataRow || e.Row.RowType == DataControlRowType.Header)
        //    {
        //        e.Row.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='#00fee';");
        //        e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor=c;");



        //    }
        //    //可以根据需要设置更多的列


        //}



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


        protected void Button2_Click(object sender, EventArgs e)
        {
            GridView1.PageIndex = 0;

        }

    }
}