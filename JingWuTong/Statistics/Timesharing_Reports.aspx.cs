using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong.Statistics
{
    public partial class Timesharing_Reports : System.Web.UI.Page
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


                this.head.InnerHtml = gethtml();


            }


        }

        public string gethtml()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<tr>");
            sb.Append("<th colspan='17' style='text-align:center'>分时段统计汇总</th>");
            sb.Append("</tr>");
            sb.Append("<tr>");
            sb.Append("<th rowspan='2' style='text-align:center'>部门</th>");
            sb.Append("<th rowspan='2' style='text-align:center'>设备配发数(台)</th>");
            sb.Append("<th colspan='3' style='text-align:center'>时间</th>");
            sb.Append("<th colspan='3' style='text-align:center'>时间</th>");
            sb.Append("<th colspan='3' style='text-align:center'>时间</th>");
            sb.Append("<th colspan='3' style='text-align:center'>时间</th>");
            sb.Append("<th colspan='3' style='text-align:center'>时间</th>");
            sb.Append("</tr>");
            sb.Append("<tr>");
            //---------------------------------------------------------
            sb.Append("<th>设备使用数量</th>");
            sb.Append("<th>在线时长总和</th>");
            sb.Append("<th>设备使用率</th>");
            //---------------------------------------------------------
            sb.Append("<th>设备使用数量</th>");
            sb.Append("<th>在线时长总和</th>");
            sb.Append("<th>设备使用率</th>");
            //---------------------------------------------------------
            sb.Append("<th>设备使用数量</th>");
            sb.Append("<th>在线时长总和</th>");
            sb.Append("<th>设备使用率</th>");
            //---------------------------------------------------------;
            sb.Append("<th>设备使用数量</th>");
            sb.Append("<th>在线时长总和</th>");
            sb.Append("<th>设备使用率</th>");
            //---------------------------------------------------------
            sb.Append("<th>设备使用数量</th>");
            sb.Append("<th>在线时长总和</th>");
            sb.Append("<th>设备使用率</th>");
            sb.Append("</tr>");

            return Convert.ToString(sb);
        
        }





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





    }
}