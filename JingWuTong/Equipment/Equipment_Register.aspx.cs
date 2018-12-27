using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong.Equipment
{
    public partial class Equipment_Register : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {
                BLL.B_Role B_Role_Load = new BLL.B_Role();
                string[] s_power = ((string)B_Role_Load.Exists(TOOL.Login.I_JSID)).Split('|');//权限参数

                if (s_power[11] == "0")
                {


                    this.B_Add.Visible = false;  //新增

                }

                if (s_power[12] == "0") //删除
                {
                    this.GridView1.Columns[11].Visible = false;
  
                    //    System.Web.UI.HtmlControls.HtmlImage Imagedel = (System.Web.UI.HtmlControls.HtmlImage)this.GridView1.Rows[i].FindControl("img_del");
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

                if (s_power[16] == "0")//导入
                {
                    this.B_Into.Visible = false;
                }

                if (s_power[17] == "0")//打印
                {
                    this.B_Print.Visible = false;
                }



                //this.T_strat.Text = DateTime.Now.AddDays(-15).ToString("yyyy/MM/dd");
                //this.T_now.Text = System.DateTime.Now.ToString("yyyy/MM/dd");

                if (HttpContext.Current.Request.Cookies["cookieName"] != null)
                {
                    HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];
                    this.L_BXR.InnerText = Server.UrlDecode(cookies["name"]);
                  
                }

                this.L_LogTime.InnerText = System.DateTime.Now.ToString("yyyy-MM-dd");

                //设备名称
                BLL.B_DeviceType dr_DeviceType = new BLL.B_DeviceType();
                this.Dr_DeviceNmae.DataSource = dr_DeviceType.GetList();
                this.Dr_DeviceNmae.DataTextField = "TypeName";
                this.Dr_DeviceNmae.DataValueField = "ID";
                this.Dr_DeviceNmae.DataBind();

                //this.Dr_DeviceNmae.Items.Insert(0, new ListItem("---请选择---", "-1"));
                this.Dr_DeviceNmae.Items.Insert(0, new ListItem("全部", "-1"));
                this.Dr_DeviceNmae.SelectedIndex = 0;
                //this.Dr_DeviceNmae.ForeColor = Color.White;
                //this.Dr_DeviceNmae.BackColor = Color.Transparent;
                //this.Dr_DeviceNmae.BorderStyle = BorderStyle.Ridge;

                //设备状态
                BLL.B_Dev_WorkState dr_Dev_WorkState = new BLL.B_Dev_WorkState();
                this.Dr_State.DataSource = dr_Dev_WorkState.GetList();
                this.Dr_State.DataTextField = "StateName";
                this.Dr_State.DataValueField = "ID";
                this.Dr_State.DataBind();
                //this.Dr_State.Items.Insert(0, new ListItem("---请选择---", "-1"));
                this.Dr_State.Items.Insert(0, new ListItem("全部", "-1"));
                this.Dr_State.SelectedIndex = 0;



                //AllocateState设备状态
                BLL.B_Dev_AllocateState dr_Dev_AllocateState = new BLL.B_Dev_AllocateState();
                this.Dr_AllocateState.DataSource = dr_Dev_AllocateState.GetList();
                this.Dr_AllocateState.DataTextField = "StateName";
                this.Dr_AllocateState.DataValueField = "ID";
                this.Dr_AllocateState.DataBind();
                //this.Dr_State.Items.Insert(0, new ListItem("---请选择---", "-1"));
                this.Dr_AllocateState.Items.Insert(0, new ListItem("全部", "-1"));
                this.Dr_AllocateState.SelectedIndex = 0;


                string s_BMDM=null;
                string s_SJBM = null;

                  if (HttpContext.Current.Request.Cookies["cookieName"] != null)
                {

                    HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];
                
                    if (cookies["BMDM"] != null)
                    {
                       s_BMDM=Server.UrlDecode(cookies["BMDM"]);
                       s_SJBM=Server.UrlDecode(cookies["SJBM"]); 
                    }
                 

                }



                //使用人单位
                BLL.Entity F_Entity = new BLL.Entity();

          

                    this.dr_first.DataSource = F_Entity.GetList("330000000000","");
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

        //protected void Page_PreRender(object sender, EventArgs e)
        //{
        //    this.dr_second.
       


        //}





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
            try
            {
               
                switch (e.CommandName)
                {
                    case "MyDel":

                        BLL.Device Del_Device = new BLL.Device();

                        string[] estr = e.CommandArgument.ToString().Split(',');

                        long ParentID = long.Parse(estr[0]);
                        string S_JYBH = estr[1].ToString();


                        bool b_JYBH = Del_Device.ExistsJYBH2(S_JYBH,ParentID);



                        if (b_JYBH == true&&S_JYBH!="")
                        {
                            Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('请将该设备与警员解除绑定，再删除设备');</script>");
                            break;
                        }

                        else
                        {

                            bool b_Device = Del_Device.Delete(ParentID);


                            if (b_Device == true)
                            {

                                //记入日志
                                TOOL.Login.OptObject = ParentID.ToString();//操作ID
                                TOOL.Login.BZ = "操作后:无";//备注

                                OperationLog OperationLogAddupdate = new OperationLog();
                                OperationLogAddupdate.OperationLogAdd("01", "05");


                                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script> alert('删除成功！');</script>");
                                this.GridView1.DataBind();
                                //this.ObjectDataSource1.DataBind();
                                //GridView1.PageIndex = 0;
                               
                            }
                            else
                            {
                                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('删除失败！');</script>");

                            }

                            break;
                        }
                    default:

                        break;
                }
            }

            catch (Exception ex)
            {
                string a = ex.Message;
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('"+a+"');</script>");
              
            }
        
        
        }




        protected void Button2_Click(object sender, EventArgs e)
        {
            this.Dr_DeviceNmae.Items[0].Value = "-2";
            this.Dr_State.Items[0].Value = "-2";
            this.Dr_AllocateState.Items[0].Value = "-2";

            GridView1.PageIndex = 0;

        }





       /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnOutExcel_Click(object sender, EventArgs e)
        {

 

          int s_DeviceNmae= int.Parse(this.Dr_DeviceNmae.SelectedValue);//设备名称

          string s_strat=this.T_strat.Text.Trim();//采购开始时间

            string s_now=this.T_now.Text.Trim();
            int i_State=int.Parse(this.Dr_State.SelectedValue);//设备状态

            string s_three=this.dr_three.SelectedValue;//使用人单位

            string s_DevId = this.T_search.Text.Trim();//设备编号模糊查询

            string s_second = this.dr_second.SelectedValue;//使用人单位


            int s_AllocateState =int.Parse( this.Dr_DeviceNmae.SelectedValue);//配发状态

            DataTable thisTable = null;

            BLL.Device B_Device = new BLL.Device();

            thisTable = B_Device.OutExcel(s_DeviceNmae, s_strat, s_now, i_State, s_three, s_second, s_DevId, s_AllocateState);

            if (thisTable.Rows.Count <= 0)
            {
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('没有数据！！！');</script>");
                return;
            
            }
       

            if (thisTable != null)
            {
                StringWriter sw = new StringWriter();
                //生成列
                sw.WriteLine("设备名称\t厂家\t设备编号\t设备型号\t规格\t项目名称\t项目编号\t单价\t项目负责人\t项目负责人联系电话\t采购时间\t保修期\t报废期限\t设备状态\t本次检验时间\t下次检验时间\t警员编号\t部门名称\t车类型\t车牌号码");
                foreach (DataRow dr in thisTable.Rows)
                {

                    //生成行
                    sw.WriteLine(dr["TypeName"] + "\t" + dr["Manufacturer"] + "\t" + dr["DevId"] + "\t" + dr["SBXH"] + "\t" + dr["SBGG"] + "\t" + dr["ProjName"] + "\t" + dr["ProjNum"] + "\t" + dr["Price"] + "\t" + dr["XMFZR"] + "\t" + dr["XMFZRDH"] + "\t" + dr["CGSJ"] + "\t" + dr["BXQ"] + "\t" + dr["BFQX"] + "\t" + dr["StateName"] + "\t" + dr["BCJYSJ"] + "\t" + dr["XCJYSJ"] + "\t" + dr["JYBH"] + "\t" + dr["BMMC"] + "\t" + dr["Cartype"] + "\t" + dr["PlateNumber"]);
                }
                sw.Close();


                //记入日志
                TOOL.Login.OptObject = "";//操作ID
                TOOL.Login.BZ = "导出了" + thisTable.Rows.Count + "条数据";//备注
                OperationLog OperationLogAddupdate = new OperationLog();
                OperationLogAddupdate.OperationLogAdd("01", "11");


                Response.Clear();

                string fileName = "设备登记(" + DateTime.Now.ToString() + ")";

                Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}.xls", fileName));

                Response.ContentEncoding = System.Text.Encoding.Default;
                Response.ContentType = "application/vnd.ms-excel";
                Response.Write(sw);
                Response.End();
            }



        }





    //    public void CreateExcel(DataTable dt,string FileName)//HttpResponse Page.Response
    //{
    //    string FileType = "application/ms-excel";
    //    Response.Clear();
    //    Response.Charset = "UTF-8";
    //    Response.Buffer = true;
    //    Response.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");
    //    Response.AppendHeader("Content-Disposition", "attachment;filename=\"" + FileName + ".xls\"");
    //    Response.ContentType = FileType;
    //    string colHeaders = string.Empty;
    //    colHeaders = "标题";
    //    Response.Output.Write(colHeaders);
    //    colHeaders = string.Empty;
    //    string ls_item = string.Empty;
    //    DataRow[] myRow = dt.Select();
    //    int cl = dt.Columns.Count;
    //    foreach (DataRow row in myRow)
    //    {
    //        int count = 0;
    //        for (int i = 0; i < cl; i++)
    //        {
    //            if (i == (cl - 1))
    //            {
    //                ls_item += row[i].ToString() + "\n";
    //            }
    //            else
    //            {
    //                if(count < 2)
    //                {
    //                    ls_item = ls_item + "" + row[i].ToString() + "\t";
    //                }
    //                else
    //                {
    //                    ls_item += row[i].ToString() + "\t";
    //                }
    //            }
    //            count++;
    //        }
    //        Response.Output.Write(ls_item);
    //        ls_item = string.Empty;
    //    }
    //    Response.Output.Flush();
    //    Response.End();
    //    }






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

        //导入
        protected void btnIntExcel_Click(object sender, EventArgs e)
        {
            try
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
                    Model.Device M_Device = new Model.Device();
                    StringBuilder strSql = new StringBuilder();

                    StringBuilder strSql2 = new StringBuilder();

                    StringBuilder strSql3 = new StringBuilder();
                    StringBuilder strSql4 = new StringBuilder();
                    StringBuilder strSql5 = new StringBuilder();

                
                    for (int i = 0; i < dr.Length; i++)
                    {
                        //前面除了你需要在建立一个“upfiles”的文件夹外，其他的都不用管了，你只需要通过下面的方式获取Excel的值，然后再将这些值用你的方式去插入到数据库里面
                        string ID = "";
                        BLL.B_DeviceType B_DeviceType_Find = new BLL.B_DeviceType();

                        object O_DevType = B_DeviceType_Find.GetListID(dr[i]["设备名称"].ToString());
                        if (O_DevType == null)
                        {
                            strSql4.Append(dr[i]["设备编号"].ToString()+",");
                            continue;

                        }
                        else
                        {

                            M_Device.DevType = (int)B_DeviceType_Find.GetListID(dr[i]["设备名称"].ToString());
                        }


                        M_Device.Manufacturer = dr[i]["厂家"].ToString();


                        BLL.Device B_Device_Find = new BLL.Device();

                        if (B_Device_Find.Exists(dr[i]["设备编号"].ToString()))//设备编号重复
                        {
                            ID = dr[i]["设备编号"].ToString();//设备编号重复进入更新状态
                            //strSql.Append(dr[i]["设备编号"].ToString());
                            //continue;
                        }

                     
                            M_Device.DevId = dr[i]["设备编号"].ToString();


                            if (string.IsNullOrEmpty(dr[i]["设备编号"].ToString()))//设备编号为空
                        {
                            strSql.Append(i + ",");
                            continue;

                        }


                        M_Device.SBXH = dr[i]["设备型号"].ToString();

                        M_Device.SBGG = dr[i]["规格"].ToString();
                        M_Device.ProjName = dr[i]["项目名称"].ToString();
                        M_Device.ProjNum = dr[i]["项目编号"].ToString();
                        M_Device.Price = dr[i]["单价"].ToString();

                        M_Device.XMFZR = dr[i]["项目负责人"].ToString();
                        M_Device.XMFZRDH = dr[i]["项目负责人联系电话"].ToString();
                        if (dr[i]["采购时间"].ToString() == "")
                        {
                            M_Device.CGSJ = null;

                        }
                        else
                        {
                            M_Device.CGSJ = Convert.ToDateTime(dr[i]["采购时间"].ToString());
                        }

                        if (dr[i]["保修期"].ToString() == "")
                        {
                            M_Device.BXQ = null;

                        }
                        else
                        {
                            M_Device.BXQ = Convert.ToDateTime(dr[i]["保修期"].ToString()); //Convert.ToDateTime 转换不了“”值，EXCEL表格中如果不填写就是""值
                        }
                        if (dr[i]["报废期限"].ToString() == "")
                        {
                            M_Device.BFQX = null;

                        }
                        else
                        {
                            M_Device.BFQX = Convert.ToDateTime(dr[i]["报废期限"].ToString());
                        }

                        BLL.B_Dev_WorkState B_Dev_WorkState_Find = new BLL.B_Dev_WorkState();

                        object O_WorkState = B_Dev_WorkState_Find.GetListID(dr[i]["设备状态"].ToString());//工作状态

                        if (O_WorkState == null) //工作状态出错
                        {

                            strSql3.Append(dr[i]["设备编号"].ToString()+",");
                            continue;
                        }

                        else
                        {
                            M_Device.WorkState = (int)B_Dev_WorkState_Find.GetListID(dr[i]["设备状态"].ToString());//工作状态
                        }



                        if (dr[i]["本次检验时间"].ToString() != "" && dr[i]["下次检验时间"].ToString() != "")
                        {

                            if (M_Device.DevType == 7 || M_Device.DevType == 8)
                            {

                                M_Device.BCJYSJ = Convert.ToDateTime(dr[i]["本次检验时间"].ToString());
                                M_Device.XCJYSJ = Convert.ToDateTime(dr[i]["下次检验时间"].ToString());

                            }
                            else
                            {
                                M_Device.BCJYSJ = null;
                                M_Device.XCJYSJ = null;
                                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('只有酒精测试仪和测速仪才有检验时间,已经帮你删除检验时间');</script>");

                            }

                        }
                        else
                        {
                            M_Device.BCJYSJ = null;
                            M_Device.XCJYSJ = null;
                        }



                        M_Device.JYBH = dr[i]["警员编号"].ToString();


                        if (!string.IsNullOrEmpty(dr[i]["部门名称"].ToString()))
                        {

                            BLL.Entity Entity_Find = new BLL.Entity();

                            object O_BMDM = Entity_Find.GetListName(dr[i]["部门名称"].ToString());

                            if (O_BMDM == null)//如果输入的部门,名称有问题。
                            {
                                strSql2.Append(dr[i]["设备编号"].ToString());
                                continue;

                            }
                            else
                            {
                                M_Device.BMDM = (string)Entity_Find.GetListName(dr[i]["部门名称"].ToString());
                            }

                        }

                        else
                        {
                            M_Device.BMDM = dr[i]["部门名称"].ToString();
                        
                        }
                        M_Device.CreatDate = System.DateTime.Now;//创建时间


                        //---新加
                        M_Device.AllocateState = 2;//配发状态

                        M_Device.Cartype = dr[i]["车类型"].ToString();
                        M_Device.PlateNumber = dr[i]["车牌号码"].ToString();

                        M_Device.IMEI = dr[i]["IMEI"].ToString();
                        M_Device.SIMID = dr[i]["SIMID"].ToString();

                        try
                        {
                            BLL.Device B_Device = new BLL.Device();
                            int count = 0;

                            if (ID != "")
                            {
                                count = B_Device.UpdateData(M_Device);//更新
                            }
                            else
                            {
                            count= B_Device.Add(M_Device);
                            }

                            if (count > 0)
                            {
                                successly++;
                            }
                            else
                            {
                                strSql5.Append(dr[i]["设备编号"].ToString()+",");
                                continue;
                            
                            }
                  

                        }
                        catch (Exception ex)
                        {
                            _Result = _Result + ex.Message + "\\n\\r";
                            Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('" + _Result.ToString() + "');</script>");
                        }
                    }//循环到此结束


                    if (strSql.Length > 0)
                    {
                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些编号的设备编号为空:" + strSql.ToString() + ",');</script>");


                    }

                    if (strSql2.Length > 0)
                    {


                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些设备编号输入的部门名称无法查询到:" + strSql2.ToString() + ",');</script>");
                    }


                    if (strSql3.Length > 0)
                    {


                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些设备编号输入的设备状态无法查询到:" + strSql3.ToString() + ",');</script>");
                    }



                    if (strSql4.Length > 0)
                    {


                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些设备编号输入的设备名称无法查询到:" + strSql4.ToString() + ",');</script>");
                    }


                    if (strSql5.Length > 0)
                    {


                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些设备编号导入时出错:" + strSql5.ToString() + "');</script>");
                    }

                    if (successly == rowsnum)
                    {
                        //记入日志
                        TOOL.Login.OptObject = "";//操作ID
                        TOOL.Login.BZ = "导入了" + successly + "条数据";//备注

                        OperationLog OperationLogAddupdate = new OperationLog();
                        OperationLogAddupdate.OperationLogAdd("01", "10");

                        string strmsg = "Excle表导入成功!";


                        System.Web.HttpContext.Current.Response.Write("<Script Language='JavaScript'>window.alert('" + strmsg + "');</script>");
                    }
                    else
                    {
                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('Excle表导入失败!" + _Result + "');</script>");
                    }
                }
            }

            catch (Exception ex)
            {
                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('" + ex.Message + "');</script>");
            
            }
        }






        //流方式下载
        protected void btnDownload_Click(object sender, EventArgs e)
        {

            string sPath = AppDomain.CurrentDomain.BaseDirectory+ "\\Excel\\设备登记.xls";

            bool b = File.Exists(sPath);
            if(!b)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('设备登记表格模版不存在，请联系管理员！');</script>");
                return;
            
            }


            string fileName = "设备登记.xls";//客户端保存的文件名


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
            string s_JYBH = "";
            TOOL.Login.BZ = "操作后:";
            if (HttpContext.Current.Request.Cookies["cookieName"] != null)
            {
                HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];
                name = Server.UrlDecode(cookies["name"]);
                s_JYBH = Server.UrlDecode(cookies["JYBH"]);
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

                    int i_DevType = int.Parse(((Label)this.GridView1.Rows[i].FindControl("L_DevType")).Text);

                    string S_DevId = ((Label)this.GridView1.Rows[i].FindControl("L_DevId")).Text;

                    string S_BMDM = ((Label)this.GridView1.Rows[i].FindControl("L_BMDM")).Text;

                    string S_WorkState = ((Label)this.GridView1.Rows[i].FindControl("L_WorkState")).Text;

                    string S_AllocateState = ((Label)this.GridView1.Rows[i].FindControl("L_AllocateState")).Text;

                    string s_State = this.HiddenNState.Text.Trim();

                    if (S_WorkState == "30" && s_State == "20")//如果设备维修状态点击了报修
                    {
                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('有项目已在维修状态无法报修');</script>");
                        return;
                    
                    }


                    if (S_WorkState == "20" && s_State=="20")
                    {

                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('有项目已在报修状态');</script>");
                        return;
                    }

                    if (S_AllocateState == "1" && s_State == "1")//回收（未配发）
                    {

                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('有项目已在回收状态');</script>");
                        return;
                    }

                    if ((S_WorkState == "10"||S_WorkState == "40") && s_State == "30") //如果点击维修按钮 单当前设备状态是正常和报废的
                    {
                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('当前项目不是报修状态无法维修');</script>");
                        return;

                    }

                    if (S_WorkState == "40" && s_State=="40")
                    {

                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('有项目已在报废状态');</script>");
                        return;
                    }


                    if (S_WorkState == "30" && s_State == "30")
                    {
                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('有项目已经在维修状态');</script>");
                        return;

                    }


                    if (S_WorkState != "30" && s_State == "10")
                    {

                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('当前项目不是维修状态无法修复');</script>");
                        return;
                    }

                    if (S_WorkState == "10" && s_State == "10")
                    {
                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('有项目已经在正常状态');</script>");
                        return;

                    }



                    //long L_ID =long.Parse( this.GridView1.Rows[i].Cells[1].Text);

                    M_Device_update.ID = L_ID;
                    if (this.HiddenState.Text == "recycle")//回收=未配发
                    {
                      

                        M_Device_update.AllocateState =int.Parse(this.HiddenNState.Text);
                        i_Device = B_Device_update.UpdateAllocateState(M_Device_update);

                      
                        TOOL.Login.OptObject += M_Device_update.ID.ToString()+",";//操作ID
                        TOOL.Login.BZ += "AllocateState:" + M_Device_update.AllocateState.ToString() + ",";//备注

                        OperationLog OperationLogAddupdate = new OperationLog();
                        OperationLogAddupdate.OperationLogAdd("01", "08");
                    }

                    else
                    {
                        if (this.HiddenState.Text == "repairs")//报修
                        {
                            
                            TOOL.Login.OptObject += M_Device_update.ID.ToString() + ",";//操作ID
                            TOOL.Login.BZ += "WorkState:" + this.HiddenNState.Text + ",";//备注
                            OperationLog OperationLogAddupdate = new OperationLog();
                            OperationLogAddupdate.OperationLogAdd("01", "06");
                        }

                        if (this.HiddenState.Text == "scrap")//报废
                        {
                           
                            TOOL.Login.OptObject += M_Device_update.ID.ToString() + ",";//操作ID
                            TOOL.Login.BZ += "WorkState:" + this.HiddenNState.Text + ",";//备注
                            OperationLog OperationLogAddupdate = new OperationLog();
                            OperationLogAddupdate.OperationLogAdd("01", "09");
                        }
                        if (this.HiddenState.Text == "maintain")//同意維修
                        {
                            TOOL.Login.OptObject += M_Device_update.ID.ToString() + ",";//操作ID
                            TOOL.Login.BZ += "WorkState:" + this.HiddenNState.Text + ",";//备注
                            OperationLog OperationLogAddupdate = new OperationLog();
                            OperationLogAddupdate.OperationLogAdd("01", "07");
                        }
                        if (this.HiddenState.Text == "normal")//确认修复
                        {
                            //TOOL.Login.OptObject += M_Device_update.ID.ToString() + ",";//操作ID
                            //TOOL.Login.BZ += "WorkState:" + this.HiddenNState.Text + ",";//备注
                            //OperationLog OperationLogAddupdate = new OperationLog();
                            //OperationLogAddupdate.OperationLogAdd("01", "09");
                        }




                        M_Device_update.WorkState = int.Parse(this.HiddenNState.Text);
                     i_Device= B_Device_update.UpdateWorkState(M_Device_update);
                    }

                   //如果报修成功就往DeviceLog表中插入报修记录

                    BLL.B_DeviceLog B_DeviceLog_Add = new BLL.B_DeviceLog();

                    Model.M_DeviceLog M_DeviceLog_Add = new Model.M_DeviceLog();

                    M_DeviceLog_Add.LogType = 1;//表示维修记录

                    M_DeviceLog_Add.DevType = i_DevType;//设备类型
                    if (this.HiddenState.Text == "recycle")//回收=未配发
                    {
                        M_DeviceLog_Add.AllocateState = int.Parse(this.HiddenNState.Text);//配发状态
                        M_DeviceLog_Add.DevState = int.Parse( S_WorkState);//只有设备状态自动获取
                    }
                    else
                    {
                        M_DeviceLog_Add.AllocateState = int.Parse(S_AllocateState);//配发状态 自动获取
                        M_DeviceLog_Add.DevState = int.Parse(this.HiddenNState.Text);//设备状态
                    }
                    M_DeviceLog_Add.LogTime =System.DateTime.Now;
                    M_DeviceLog_Add.BXR = name;
                    M_DeviceLog_Add.BZ = this.T_BZ.Text.Trim();
                    M_DeviceLog_Add.DevId = S_DevId;
                    M_DeviceLog_Add.Entity = S_BMDM;
                    M_DeviceLog_Add.JYBH = s_JYBH;
 
                  i_DeviceLog= B_DeviceLog_Add.AddThree(M_DeviceLog_Add);


                }

            }

            if (i_Device > 0 && i_DeviceLog > 0)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('操作成功');</script>");
                this.GridView1.DataBind();
                //Response.Redirect("Equipment_Register.aspx");
                return;
            }



          if(I_sum<=0)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('请先选择要操作的行！！！');</script>");
                return;
            }



        }






    }
}