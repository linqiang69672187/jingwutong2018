using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.DataVisualization.Charting;
using System.Web.UI.WebControls;
using System.Web.UI.DataVisualization.Charting.Utilities;
using System.Text;

namespace JingWuTong
{
    public partial class Equipment_Examine : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string s_BMDM = "";

            if (HttpContext.Current.Request.Cookies["cookieName"] != null)
            {

                HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];

                if (cookies["BMDM"] != null)
                {
                    s_BMDM = Server.UrlDecode(cookies["BMDM"]);
                }

            }


            BLL.Entity c_Entity = new BLL.Entity();

            this.trianglebox.InnerHtml = "<table id='t_eq' align='center'  >" +
         "<caption style='color:blue;font-size:150%'>设备总数:" + c_Entity.GetAllCount(s_BMDM) + "</caption>" +
         "<tr>"+
       "<td>车载视频:" + c_Entity.GetAllDevTyCount(1, s_BMDM) + "</td> " +
       "<td>对讲机:" + c_Entity.GetAllDevTyCount(2, s_BMDM) + "</td> " +
       "<td>拦截仪:" + c_Entity.GetAllDevTyCount(3, s_BMDM) + "</td> " +
        "</tr> "+
        "<tr> "+
       "<td>执法记录仪:" + c_Entity.GetAllDevTyCount(5, s_BMDM) + "</td> " +
       "<td>辅警通:" + c_Entity.GetAllDevTyCount(6, s_BMDM) + "</td> " +
       "<td>警务通:" + c_Entity.GetAllDevTyCount(4, s_BMDM) + "</td> " +
       "</tr> "+
 
     "</table>";
          

            //ViewState["show"] = "hide";

            if (!Page.IsPostBack)
            {
                //string strWidth = Request.Form["strWidth"];

                //string strHeight = Request.Form["strHeight"];


            

                BLL.B_Role B_Role_Load = new BLL.B_Role();
                string[] s_power = ((string)B_Role_Load.Exists(TOOL.Login.I_JSID)).Split('|');//权限参数

                StringBuilder strjs = new StringBuilder();
                if (s_power[3] == "0")
                {

                    strjs.Append(" show();"); //首页 操作 0 为不可操作

                }

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>" + strjs.ToString() + "</script>");

                this.T_first.InnerText = "总设备数" + i_compute("331002000000", "331002000000");
                this.T_Second.InnerText = "总设备数" + i_compute("331003000000", "331003000000");
                this.T_Three.InnerText = "总设备数" + i_compute("331004000000", "331004000000");
                this.T_Four.InnerText = "总设备数" + i_compute("331001000000", "331001000000");

                ViewState["1"] = "0";
                ViewState["2"] = "0";
                ViewState["3"] = "0";
                ViewState["4"] = "0";
                ViewState["5"] = "0";
                ViewState["6"] = "0";
                //ViewState["name"] = "331003000000";



                //loader("331003000000", "331003000000");
                hide();

                OperationLog OperationLogAddupdate = new OperationLog();
                OperationLogAddupdate.OperationLogAdd("17", "01");

              
            }

        }

        protected void Page_PreRender(object sender, EventArgs e)
        {
            try
            {
                string s_width = (int.Parse(TOOL.Login.strWidth) *2/3).ToString();

                string s_Height = (int.Parse(TOOL.Login.strHeight)/2).ToString();

                this.Chart1.Width = Unit.Parse(s_width);

                this.Chart1.Height = Unit.Parse(s_Height);
            }
            catch (Exception ex)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>'"+ex.Message+"'</script>");
            }

        }


    






        //计算单位设备总数
        protected int i_compute(string s_SJBM, string s_BMDM)
        {


            BLL.Entity i_Entity = new BLL.Entity();
            return i_Entity.GetCount(s_SJBM, s_BMDM);


        
        }


        //绑定柱形数据
        protected void loader(string Bjbm, string Bmdm)
        {

            BLL.Entity Entity = new BLL.Entity();

            this.Chart1.Width = Unit.Parse(this.HiddenState.Text.Trim());

            this.Chart1.Height = Unit.Parse(this.HiddenNState.Text.Trim());

            this.Chart1.ChartAreas[0].AxisX.LabelStyle.ForeColor = Color.White;
            this.Chart1.ChartAreas[0].AxisY.LabelStyle.ForeColor = Color.White;

            this.Chart1.ChartAreas[0].AxisY.LineColor = Color.White;
            this.Chart1.ChartAreas[0].AxisX.LineColor = Color.White;

            this.Chart1.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.Transparent;
            this.Chart1.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.White;

            this.Chart1.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash;


            this.Chart1.ChartAreas[0].AxisX.LabelStyle.Interval = 1;


            this.Chart1.ChartAreas[0].AxisX.TitleFont=new Font("宋体",10);

        
            //this.Chart1.ChartAreas[0].AxisX.Interval = 6;
   
            //Chart1.Series[4].Label = "#VAL";
            //Chart1.Series[4].ToolTip = "XXXXX:#VAL";
            //Chart1.Series[4].Name = "123";

        
            for (int i = 0; i <= 5; i++)
            {

                switch (i + 1)
                {
                    case 1://车载视频

                        //if (Session["1"].ToString() != "1")
                        //{




                            Series series1 = new Series();
                            //series1.ChartType = SeriesChartType.StackedColumn;//一柱多列
                            series1.CustomProperties = "PointWidth=0.2";
                            series1.Color = Color.FromArgb(104, 151, 235);
                            series1.LegendText = "车载视频";
                            series1.LegendPostBackValue = "1";
             
                      
                            if (ViewState["1"].ToString() == "1")
                            {
                                series1.Enabled = false;
                            }
                            this.Chart1.Series.Add(series1);
                    
                        //}

                        break;
                    case 2://对讲机
                      
                            Series series2 = new Series();
                            //series2.ChartType = SeriesChartType.StackedColumn;//一柱多列
                            series2.CustomProperties = "PointWidth=0.2";
                            series2.Color = Color.FromArgb(235, 194, 115);
                            series2.LegendText = "对讲机";
                            series2.LegendPostBackValue = "2";
                       

                            if (ViewState["2"].ToString() == "2")
                            {
                                series2.Enabled = false;
                            }

                            this.Chart1.Series.Add(series2);
                     
                        break;
                    case 3://拦截仪
                        Series series3 = new Series();
                        //series3.ChartType = SeriesChartType.StackedColumn;//一柱多列
                        series3.CustomProperties = "PointWidth=0.2";
                        series3.Color = Color.FromArgb(84, 225, 149);
                       
              
                        series3.LegendText = "拦截仪";
                        series3.LegendPostBackValue = "3";
                        if (ViewState["3"].ToString() == "3")
                        {
                            series3.Enabled = false;
                        }

                        this.Chart1.Series.Add(series3);
                        break;
                    case 4://警务通
                        Series series4 = new Series();
                        //series4.ChartType = SeriesChartType.StackedColumn;//一柱多列
                        series4.CustomProperties = "PointWidth=0.2";
                        series4.Color = Color.FromArgb(80, 227, 227);

                        series4.LegendText = "警务通";
                        series4.LegendPostBackValue = "4";

                        if (ViewState["4"].ToString() == "4")
                        {
                            series4.Enabled = false;
                        }

                        this.Chart1.Series.Add(series4);
                        break;
                    case 5://执法记录仪
                        Series series5 = new Series();
                        //series5.ChartType = SeriesChartType.StackedColumn;//一柱多列
                        series5.CustomProperties = "PointWidth=0.2";
                        series5.Color = Color.FromArgb(231, 102, 103);
                        series5.LegendText = "执法记录仪";
                        series5.LegendPostBackValue = "5";

                        if (ViewState["5"].ToString() == "5")
                        {
                            series5.Enabled = false;
                        }

                        this.Chart1.Series.Add(series5);

                        break;
                    case 6://辅警通
                        Series series6 = new Series();
                        //series6.ChartType = SeriesChartType.StackedColumn;//一柱多列
                        series6.CustomProperties = "PointWidth=0.2";
                        series6.Color = Color.FromArgb(231, 214, 120);

                        series6.LegendText = "辅警通";
                        series6.LegendPostBackValue = "6";


                        if (ViewState["6"].ToString() == "6")
                        {
                            series6.Enabled = false;
                        }

                        this.Chart1.Series.Add(series6);


                        break;
                    default:

                        break;
                }

               


                ////设置右侧图标
                //Legend legend = new Legend();

                //this.Chart1.Legends.Add(legend);

              // this.Chart1.Legends[0].CellColumns["series1"].Name = "dsd";

           



                DataTable dt = Entity.GetEntity(Bjbm, i + 1, Bmdm);


                if (dt.Rows.Count == 0)
                {
                    return;
                }
       
                   
                


                int [] arrayY = new int[dt.Rows.Count];

                string[] arrayX = new string[dt.Rows.Count];
                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    DataRow dr2 = dt.Rows[x];
                    arrayY[x] = int.Parse(dr2["DevType"].ToString());

                    arrayX[x] = Convert.ToString(dr2["BMMC"]);

                    //if (int.Parse(dr2["DevType"].ToString()) != 0)
                    //{
                    //    this.Chart1.Series[0].
                    //}

                }



                this.Chart1.Series[i].ChartType = SeriesChartType.StackedColumn;
                //this.Chart1.Series[i].IsValueShownAsLabel = true;
                //this.Chart1.Series[i].LabelForeColor = Color.Transparent;
                //this.Chart1.Series[i].Color = Color.Transparent;
                //this.Chart1.Series[i].LabelFormat = " #VALY";
                //this.Chart1.Series[i].LabelToolTip = "#VALX: #VALY";
                this.Chart1.Series[i].ToolTip = "#VALX: #VALY";
                
             

                //if (ViewState["show"].ToString() == "show")
                //{
                //    this.Chart1.Series[i].Enabled = true;
                
                //}

         
                //this.Chart1.DataSource = dt;
               // this.Chart1.Series[i].XValueMember = "BMMC";//X轴数据成员列
               // this.Chart1.Series[i].YValueMembers = "DevType";//Y轴数据成员列

                this.Chart1.Series[i].Points.DataBindXY(arrayX, arrayY);

 

         
          

            }

         
          

        
        }




        /// <summary>
        ///权限设置
        /// </summary>
        public void hide()
        {
            string s_BMDM="";

            if (HttpContext.Current.Request.Cookies["cookieName"] != null)
            {

                HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];

                if (cookies["BMDM"] != null)
                {
                    s_BMDM = Server.UrlDecode(cookies["BMDM"]);
                }
       
            }

            if (s_BMDM == "331000000000")//支队
            {
                this.one.Style.Add("visibility", "inherit");
                this.two.Style.Add("visibility", "inherit");
                this.three.Style.Add("visibility", "inherit");
                this.four.Style.Add("visibility", "inherit");

                ViewState["name"] = "331002000000";
                loader("331002000000", "331002000000");

                Page.ClientScript.RegisterStartupScript(GetType(), "", "<script>C_div('one');</script>");


            }
            else
            {
                BLL.Entity B_Entity_Find = new BLL.Entity();
             string s_SJBM=(string)B_Entity_Find.ExistsSJBM(s_BMDM);

             if (s_SJBM == "331000000000")
             {
                 switch (s_BMDM)
                 {
                     case "331002000000"://直属一大队
                         this.one.Style.Add("visibility", "inherit");
                   ViewState["name"] = "331002000000";
                   loader("331002000000", "331002000000");

                   Page.ClientScript.RegisterStartupScript(GetType(), "", "<script>C_div('one');</script>");
                         break;
                     case "331003000000":  //直属二大队
                         this.two.Style.Add("visibility", "inherit");
                ViewState["name"] = "331003000000";
                loader("331003000000", "331003000000");
                Page.ClientScript.RegisterStartupScript(GetType(), "", "<script>C_div('two');</script>");
                         break;
                     case "331004000000": //直属三大队
                         this.three.Style.Add("visibility", "inherit");

             ViewState["name"] = "331004000000";
            loader("331004000000", "331004000000");
            Page.ClientScript.RegisterStartupScript(GetType(), "", "<script>C_div('three');</script>");
                         break;
                     case "331001000000":  //直属四大队
                         this.four.Style.Add("visibility", "inherit");

             ViewState["name"] = "331001000000";
             loader("331001000000", "331001000000");
             Page.ClientScript.RegisterStartupScript(GetType(), "", "<script>C_div('four');</script>");
                         break;


                     default:

                         break;
                 }

             }
             else//中队
             {

                 ViewState["name"] = s_BMDM;

                 loader(s_BMDM, s_BMDM);

             
             }


             
            }

        
        }




        //直属一大队
        protected void Button1_Click(object sender, EventArgs e)
        {
            //ViewState["show"] = "show";
            ViewState["name"] = "331002000000";

            loader("331002000000", "331002000000");


            Page.ClientScript.RegisterStartupScript(GetType(), "", "<script>C_div('one');</script>");

        }

        //直属二大队
        protected void Button2_Click(object sender, EventArgs e)
        {

            //ViewState["show"] = "show";
            ViewState["name"] = "331003000000";

            loader("331003000000", "331003000000");
            Page.ClientScript.RegisterStartupScript(GetType(), "", "<script>C_div('two');</script>");

        }

        //直属三大队
        protected void Button3_Click(object sender, EventArgs e)
        {
            //ViewState["show"] = "show";

            ViewState["name"] = "331004000000";
            loader("331004000000", "331004000000");
            Page.ClientScript.RegisterStartupScript(GetType(), "", "<script>C_div('three');</script>");

        }

        //直属四大队
        protected void Button4_Click(object sender, EventArgs e)
        {
            //ViewState["show"] = "show";
            ViewState["name"] = "331001000000";

            loader("331001000000", "331001000000");
            Page.ClientScript.RegisterStartupScript(GetType(), "", "<script>C_div('four');</script>");

        }



        //全部显示
        protected void BShow_Click()
        {

            ViewState["1"] = "0";
            ViewState["2"] = "0";
            ViewState["3"] = "0";
            ViewState["4"] = "0";
            ViewState["5"] = "0";
            ViewState["6"] = "0";

           // ViewState["show"] = "show";
            loader(ViewState["name"].ToString(), ViewState["name"].ToString());
        }


       

   
        /// <summary>
        /// chart 点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Chart1_Click(object sender, ImageMapEventArgs e)
        {
           
            //loader( ViewState["name"].ToString(), 0);
            string s_values = e.PostBackValue;
            //if (s_values == "1")
            //{
            //    ViewState["1"] = "1";
            //    //this.Chart1.Series.RemoveAt(0);
            //}

            //if (s_values == "2")
            //{
            //    ViewState["2"] = "2";

            //    //this.Chart1.Series.RemoveAt(1);
            //}

            if (s_values == "true")
            {
                BShow_Click();

            }
            else
            {

                switch (s_values)
                {
                    case "1":
                        ViewState["1"] = "1";
                        break;
                    case "2":
                        ViewState["2"] = "2";
                        break;
                    case "3":
                        ViewState["3"] = "3";
                        break;
                    case "4":
                        ViewState["4"] = "4";
                        break;
                    case "5":
                        ViewState["5"] = "5";
                        break;
                    case "6":
                        ViewState["6"] = "6";
                        break;
                    default:
                        break;
                }


                loader(ViewState["name"].ToString(), ViewState["name"].ToString());
            }
        }

        protected void Chart1_CustomizeLegend(object sender, CustomizeLegendEventArgs e)
        {

          
   
            e.LegendItems.Add(Color.Red,"二维");


        }

     




    }
}