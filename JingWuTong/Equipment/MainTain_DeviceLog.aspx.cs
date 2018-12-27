using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong.Maintain
{
    public partial class MainTain_DeviceLog : System.Web.UI.Page
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
                    this.GridView1.Columns[13].Visible = false;

                }

                if (s_power[13] == "0")//编辑
                {
                    this.GridView1.Columns[12].Visible = false;
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


                //设备状态
                BLL.B_Dev_WorkState dr_Dev_WorkState = new BLL.B_Dev_WorkState();
                this.Dr_State.DataSource = dr_Dev_WorkState.GetListrEpairs();
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
                    this.L_BXR.InnerText = Server.UrlDecode(cookies["name"]);

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


                this.L_LogTime.InnerText = System.DateTime.Now.ToString("yyyy-MM-dd");


     

            }

        }

        //所属部门 二级
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
                        int i_DevState = int.Parse(estr[1]);

                        if (i_DevState != 20)
                        {
                            Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('只有报修中的设备可以允许删除！');</script>");
                            break ;
                        }




                        bool b_ID = B_DeviceLog_Del.Delete(l_ParentID);

                     if (b_ID == true)
                   {
                       //记入日志
                       TOOL.Login.OptObject = l_ParentID.ToString();//操作ID
                       TOOL.Login.BZ = "操作后:无";//备注
                       OperationLog OperationLogAddupdate = new OperationLog();
                       OperationLogAddupdate.OperationLogAdd("02", "05");

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
            this.Dr_DeviceNmae.Items[0].Value = "-2";
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

            int i_DeviceNmae = int.Parse(this.Dr_DeviceNmae.SelectedValue);//设备名称

            string s_strat = this.T_strat.Text.Trim();//采购开始时间
            string s_now = this.T_now.Text.Trim();


            int i_State = int.Parse(this.Dr_State.SelectedValue);//设备状态

            string s_three = this.dr_three.SelectedValue;//所属部门

            string s_search = this.T_search.Text.Trim();

            string s_second = this.dr_second.SelectedValue;

            //int i_Index = this.GridView1.PageIndex * 10;

            //int i_Count = 10;


            DataTable thisTable = null;


            BLL.B_DeviceLog B_DeviceLog_Out = new BLL.B_DeviceLog();

            thisTable = B_DeviceLog_Out.OutExcel(s_three,s_second, i_DeviceNmae, s_strat, s_now, i_State, s_search);

            if (thisTable.Rows.Count <= 0)
            {
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('没有数据！！！');</script>");
                return;

            }



            if (thisTable != null)
            {
                StringWriter sw = new StringWriter();
                //生成列
                sw.WriteLine("日志类型\t所属部门\t报修人\t联系电话\t警号\t设备名称\t设备状态\t设备编号\t申请时间\t备注");
                foreach (DataRow dr in thisTable.Rows)
                {

                    //生成行
                    sw.WriteLine(dr["LogType"] + "\t" + dr["BMMC"] + "\t" + dr["BXR"] + "\t" + dr["Tel"] + "\t" + dr["JYBH"] + "\t" + dr["TypeName"] + "\t" + dr["StateName"] + "\t" + dr["DevId"] + "\t" + dr["LogTime"] + "\t" + dr["BZ"]);
                }
                sw.Close();

                //记入日志
                TOOL.Login.OptObject = "";//操作ID
                TOOL.Login.BZ = "导出了" + thisTable.Rows.Count + "条数据";//备注
                OperationLog OperationLogAddupdate = new OperationLog();
                OperationLogAddupdate.OperationLogAdd("02", "11");

                string fileName = "维修统计(" + DateTime.Now.ToString("yyyy-MM-dd") + ").xls";

                Response.AddHeader("Content-Disposition", "attachment;  filename=" + HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8));

                Response.ContentType = "application/ms-excel";
                Response.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");
                Response.Write(sw);
                Response.End();
            }


      
        }











        //导入
        public System.Data.DataTable GetExcelDatatable(string fileUrl)
        {
            //支持.xls和.xlsx，即包括office2010等版本的   HDR=Yes代表第一行是标题，不是数据；
            string cmdText = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'";
            System.Data.DataTable dt = null;
            //建立连接
            OleDbConnection conn = new OleDbConnection(string.Format(cmdText, fileUrl));
            try
            {
                //打开连接
                if (conn.State == ConnectionState.Broken || conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }

                System.Data.DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string strSql = "select * from [Sheet1$]";
                OleDbDataAdapter da = new OleDbDataAdapter(strSql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds);
                dt = ds.Tables[0];
                return dt;
            }
            catch (Exception exc)
            {
                throw exc;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
        }

        protected void btnIntExcel_Click(object sender, EventArgs e)
        {
            if (FileUpload1.HasFile == false)//HasFile用来检查FileUpload是否有指定文件
            {
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('请您选择Excel文件');</script>");

                return;//当无文件时,返回
            }
            string IsXls = Path.GetExtension(FileUpload1.FileName).ToString().ToLower();//System.IO.Path.GetExtension获得文件的扩展名
            if (IsXls != ".xlsx" && IsXls != ".xls")
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('只可以选择Excel文件');</script>");
                return;//当选择的不是Excel文件时,返回
            }



            string filename = FileUpload1.FileName;              //获取Execle文件名  DateTime日期函数



            string savePath = AppDomain.CurrentDomain.BaseDirectory + "uploadfiles\\" + filename;//Server.MapPath 获得虚拟服务器相对路径
            DataTable ds = new DataTable();
            FileUpload1.SaveAs(savePath);                        //SaveAs 将上传的文件内容保存在服务器上
            ds = GetExcelDatatable(savePath);           //调用自定义方法
            DataRow[] dr = ds.Select();            //定义一个DataRow数组
            int rowsnum = ds.Rows.Count;
            int successly = 0;
            if (rowsnum == 0)
            {
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('Excel表为空表,无数据!');</script>"); //当Excel表为空时,对用户进行提示
            }
            else
            {
                string _Result = "";

                Model.M_DeviceLog M_DeviceLog_Int = new Model.M_DeviceLog();
                for (int i = 0; i < dr.Length; i++)
                {

                    //前面除了你需要在建立一个“upfiles”的文件夹外，其他的都不用管了，你只需要通过下面的方式获取Excel的值，然后再将这些值用你的方式去插入到数据库里面

                    M_DeviceLog_Int.LogType = int.Parse(dr[i]["日志类型"].ToString());//1:表示维修记录；2:表示设备日志

                    M_DeviceLog_Int.DevType =int.Parse(dr[i]["设备类型"].ToString());
                    M_DeviceLog_Int.DevId = dr[i]["设备编号"].ToString();
                    M_DeviceLog_Int.DevState = int.Parse( dr[i]["设备状态"].ToString());
                    M_DeviceLog_Int.JYBH = dr[i]["警员编号"].ToString();

                    M_DeviceLog_Int.BXR = dr[i]["报修人"].ToString();
                    M_DeviceLog_Int.Entity = dr[i]["用户名称"].ToString();

                    M_DeviceLog_Int.Tel = dr[i]["电话"].ToString();
                    M_DeviceLog_Int.BZ = dr[i]["备注"].ToString();
                    M_DeviceLog_Int.Entity = dr[i]["单位名称"].ToString();
                    M_DeviceLog_Int.LogTime =Convert.ToDateTime( dr[i]["日志时间"].ToString());

                


                    try
                    {
                        BLL.B_DeviceLog B_DeviceLog_Int = new BLL.B_DeviceLog();
                        int count = B_DeviceLog_Int.Add(M_DeviceLog_Int);
                        if (count > 0)
                            successly++;

                    }
                    catch (Exception ex)
                    {
                        _Result = _Result + ex.InnerException + "\\n\\r";
                    }
                }
                if (successly == rowsnum)
                {

                    OperationLog OperationLogAddupdate = new OperationLog();
                    OperationLogAddupdate.OperationLogAdd("02", "10");
                    string strmsg = "Excle表导入成功!";


                    System.Web.HttpContext.Current.Response.Write("<Script Language='JavaScript'>window.alert('" + strmsg + "');</script>");
                }
                else
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('Excle表导入失败!');</script>");
                }
            }
        }







        //流方式下载
        protected void btnDownload_Click(object sender, EventArgs e)
        {

            string sPath = AppDomain.CurrentDomain.BaseDirectory + "\\Excel\\维修统计.xls";

            bool b = File.Exists(sPath);
            if (!b)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('设备登记表格模版不存在，请联系管理员！');</script>");
                return;

            }


            string fileName = "维修统计.xls";//客户端保存的文件名


            //以字符流的形式下载文件
            FileStream fs = new FileStream(sPath, FileMode.Open);
            byte[] bytes = new byte[(int)fs.Length];
            fs.Read(bytes, 0, bytes.Length);
            fs.Close();
            Response.ContentType = "application/octet-stream";
            //通知浏览器下载文件而不是打开
            Response.AddHeader("Content-Disposition", "attachment;  filename=" + HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8));
            Response.BinaryWrite(bytes);
            Response.Flush();
            Response.End();
        }



        //一键公用功能
        protected void btnScrap_Click(object sender, EventArgs e)
        {

            string name = "";
            if (HttpContext.Current.Request.Cookies["cookieName"] != null)
            {
                HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];
                name =Server.UrlDecode(cookies["name"]);

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

                    // 根据ID更新DeviceLog 中4个字段

                    long L_ID = long.Parse(((Label)this.GridView1.Rows[i].FindControl("L_ID")).Text);

                    string S_DevState = ((Label)this.GridView1.Rows[i].FindControl("L_DevState")).Text;


                    string s_State = this.HiddenNState.Text.Trim();

                    if (S_DevState == "30" && s_State=="30")
                    {
                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('有项目已经在维修状态');</script>");
                        return;
                    
                    }

                    if (S_DevState == "10" && s_State=="10")
                    {
                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('有项目已经在正常状态');</script>");
                        return;

                    }

                    BLL.B_DeviceLog B_DeviceLog_update= new BLL.B_DeviceLog();

                    Model.M_DeviceLog M_DeviceLog_update = new Model.M_DeviceLog();

                    M_DeviceLog_update.ID = L_ID;
                    M_DeviceLog_update.LogTime = Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd"));
                    M_DeviceLog_update.BXR = name;
                    M_DeviceLog_update.BZ = this.T_BZ.Text.Trim();
                    M_DeviceLog_update.DevState = int.Parse(this.HiddenNState.Text);//设备状态

                   i_Device = B_DeviceLog_update.UpdateFour(M_DeviceLog_update);





                    //更新Device表中WorkState 工作状态字段
                    BLL.Device B_Device_update = new BLL.Device();
                    Model.Device M_Device_update = new Model.Device();

                    M_Device_update.DevId= this.GridView1.Rows[i].Cells[8].Text;

                    M_Device_update.WorkState = int.Parse(this.HiddenNState.Text);

                  i_DeviceLog=  B_Device_update.UpdateWorkState_DevId(M_Device_update);



                  if (this.HiddenState.Text == "maintain")//同意维修 日志
                  {

                      OperationLog OperationLogAddupdate = new OperationLog();
                      OperationLogAddupdate.OperationLogAdd("02", "07");
                  }


                  //if (this.HiddenState.Text == "normal")//确认修复
                  //{

                  //    OperationLog OperationLogAddupdate = new OperationLog();
                  //    OperationLogAddupdate.OperationLogAdd("02", "07");
                  //}



                }

            }

            if (i_Device > 0 && i_DeviceLog > 0)
            {
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('操作成功');</script>");
                this.GridView1.DataBind();
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