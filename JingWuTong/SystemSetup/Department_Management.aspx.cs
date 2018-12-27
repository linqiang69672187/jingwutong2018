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

namespace JingWuTong.SystemSetup
{
    public partial class Department_Management : BasePage
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {

                BLL.B_Role B_Role_Load = new BLL.B_Role();
                string[] s_power = ((string)B_Role_Load.Exists(TOOL.Login.I_JSID)).Split('|');//权限参数

                //if (s_power[25] == "0")
                //{


                //    this.B_Add.Visible = false;  //新增

                //}

                //if (s_power[26] == "0") //删除
                //{
                //    this.GridView1.Columns[11].Visible = false;

                //}

                if (s_power[27] == "0")//编辑
                {
                    this.GridView1.Columns[12].Visible = false;
                }

                if (s_power[28] == "0")//搜索
                {
                    this.Button1.Visible = false;
                }

        

                if (s_power[29] == "0")//导入
                {
                    this.B_Into.Visible = false;
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



                this.dr_first.DataSource = F_Entity.GetList2("330000000000", "");
                this.dr_first.DataTextField = "BMMC";
                this.dr_first.DataValueField = "BMDM";
                this.dr_first.DataBind();
                this.dr_first.SelectedIndex = 0;


                if (s_SJBM == "330000000000")
                {
                    this.dr_second.DataSource = F_Entity.GetList2(this.dr_first.SelectedValue, "");
                    this.dr_second.DataTextField = "BMMC";
                    this.dr_second.DataValueField = "BMDM";
                    this.dr_second.DataBind();
                    this.dr_second.Items.Insert(0, new ListItem("全部", ""));
                    this.dr_second.SelectedIndex = 0;

                    this.dr_three.DataSource = F_Entity.GetList2(this.dr_second.SelectedValue, "");
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

                        this.dr_second.DataSource = F_Entity.GetList2(this.dr_first.SelectedValue, s_SJBM);
                        this.dr_second.DataTextField = "BMMC";
                        this.dr_second.DataValueField = "BMDM";
                        this.dr_second.DataBind();
                        this.dr_second.SelectedIndex = 0;

                        this.dr_three.DataSource = F_Entity.GetList2(this.dr_second.SelectedValue, s_BMDM);
                        this.dr_three.DataTextField = "BMMC";
                        this.dr_three.DataValueField = "BMDM";
                        this.dr_three.DataBind();
                        this.dr_three.SelectedIndex = 0;
                    }
                    else
                    {
                        this.dr_second.DataSource = F_Entity.GetList2(this.dr_first.SelectedValue, s_BMDM);
                        this.dr_second.DataTextField = "BMMC";
                        this.dr_second.DataValueField = "BMDM";
                        this.dr_second.DataBind();
                        this.dr_second.SelectedIndex = 0;



                        this.dr_three.DataSource = F_Entity.GetList2(this.dr_second.SelectedValue, "");
                        this.dr_three.DataTextField = "BMMC";
                        this.dr_three.DataValueField = "BMDM";
                        this.dr_three.DataBind();
                        this.dr_three.Items.Insert(0, new ListItem("全部", ""));
                        this.dr_three.SelectedIndex = 0;

                    }

                }
            }

        }




          //同步
        protected void Btn_Click(object sender, EventArgs e)
        {
            bool syncDepartResult = SyncDepart();
            if (syncDepartResult)
            {
                Response.Write("<script>alert('部门同步成功');location.href='Department_Management.aspx'</script>");
                Response.End();
            }
            else
            {
                Response.Write("<script>alert('部门同步失败');location.href='Department_Management.aspx'</script>");
                Response.End();
            }


        }


        protected void dr_second_SelectedIndexChanged(object sender, EventArgs e)
        {

            BLL.Entity F_Entity = new BLL.Entity();

            DataTable ds_Source = F_Entity.GetList2(this.dr_second.SelectedValue, "");

            this.dr_three.DataSource = F_Entity.GetList2(this.dr_second.SelectedValue, "");
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





        //更新前
        protected void GridView1_RowEditing(object sender, GridViewEditEventArgs e)
        {



     

            GridView1.EditIndex = e.NewEditIndex;



        }


        //更新
        protected void GridView1_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {

            try
            {


                Model.Entity M_Entity_Update = new Model.Entity();



                //M_Entity_Update.BMMC = ((TextBox)(GridView1.Rows[e.RowIndex].Cells[2].FindControl("T_BM"))).Text.ToString().Trim();

                M_Entity_Update.BMJC = ((TextBox)(GridView1.Rows[e.RowIndex].Cells[3].FindControl("T_BMJC"))).Text.ToString().Trim();

                //BLL.Entity B_Entity_Update = new BLL.Entity();

                //if (B_Entity_Update.GetListBMDM(((TextBox)(GridView1.Rows[e.RowIndex].Cells[4].FindControl("T_SJ"))).Text.ToString().Trim()) == null)
                //{
                //    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('没有查询到对应上级部门！');</script>");

                //}

                //M_Entity_Update.SJBM = (string)B_Entity_Update.GetListBMDM(((TextBox)(GridView1.Rows[e.RowIndex].Cells[4].FindControl("T_SJ"))).Text.ToString().Trim());//上级单位

                M_Entity_Update.LXDZ = ((TextBox)(GridView1.Rows[e.RowIndex].Cells[5].FindControl("T_LXDZ"))).Text.ToString().Trim();




                M_Entity_Update.Lo = decimal.Parse(((TextBox)(GridView1.Rows[e.RowIndex].Cells[6].FindControl("T_Lo"))).Text.ToString().Trim() == "" ? "0" : ((TextBox)(GridView1.Rows[e.RowIndex].Cells[6].FindControl("T_Lo"))).Text.ToString().Trim());

                M_Entity_Update.La = decimal.Parse(((TextBox)(GridView1.Rows[e.RowIndex].Cells[6].FindControl("T_La"))).Text.ToString().Trim() == "" ? "0" : ((TextBox)(GridView1.Rows[e.RowIndex].Cells[6].FindControl("T_La"))).Text.ToString().Trim()); //经度纬度



                //M_Entity_Update.BMDM = ((TextBox)(GridView1.Rows[e.RowIndex].Cells[7].FindControl("T_BMDM"))).Text.ToString().Trim();
                M_Entity_Update.FZR = ((TextBox)(GridView1.Rows[e.RowIndex].Cells[8].FindControl("T_FZR"))).Text.ToString().Trim();


                M_Entity_Update.LXDH = ((TextBox)(GridView1.Rows[e.RowIndex].Cells[9].FindControl("T_LXDH"))).Text.ToString().Trim();

                M_Entity_Update.Sort = int.Parse(((TextBox)(GridView1.Rows[e.RowIndex].Cells[10].FindControl("T_Sort"))).Text.ToString().Trim() == "" ? "0" : ((TextBox)(GridView1.Rows[e.RowIndex].Cells[10].FindControl("T_Sort"))).Text.ToString().Trim());//排序



                M_Entity_Update.FY = ((TextBox)(GridView1.Rows[e.RowIndex].Cells[11].FindControl("T_FY"))).Text.ToString().Trim();

                M_Entity_Update.IsDel = int.Parse(((DropDownList)(GridView1.Rows[e.RowIndex].Cells[12].FindControl("D_IsDel"))).Text.ToString().Trim());//是否显示


                M_Entity_Update.GXSJ = System.DateTime.Now;//更新时间

                M_Entity_Update.ID = int.Parse(((Label)(GridView1.Rows[e.RowIndex].Cells[1].FindControl("L_ID"))).Text.ToString().Trim());
                M_Entity_Update.BMDM = ((Label)(GridView1.Rows[e.RowIndex].Cells[1].FindControl("L_BMDM"))).Text.ToString().Trim();

                //TOOL.Login.BZ = "操作前:" + M_Entity_Update.Statistics();//备注
     
                //BLL.Device B_Device_update = new BLL.Device();
                //B_Device_update.UpdateRemind(M_Device_update);



                this.ObjectDataSource1.UpdateMethod = "UpdateDepartment";

                this.ObjectDataSource1.UpdateParameters.Add("BMMC", DbType.String, "");
                this.ObjectDataSource1.UpdateParameters.Add("BMJC", DbType.String, M_Entity_Update.BMJC);
                this.ObjectDataSource1.UpdateParameters.Add("SJBM", DbType.String, "");
                this.ObjectDataSource1.UpdateParameters.Add("LXDZ", DbType.String, M_Entity_Update.LXDZ);
                this.ObjectDataSource1.UpdateParameters.Add("Lo", DbType.Decimal, M_Entity_Update.Lo.ToString());
                this.ObjectDataSource1.UpdateParameters.Add("La", DbType.Decimal, M_Entity_Update.La.ToString());
                this.ObjectDataSource1.UpdateParameters.Add("BMDM", DbType.String, M_Entity_Update.BMDM);
                this.ObjectDataSource1.UpdateParameters.Add("FZR", DbType.String, M_Entity_Update.FZR);
                this.ObjectDataSource1.UpdateParameters.Add("LXDH", DbType.String, M_Entity_Update.LXDH);
                this.ObjectDataSource1.UpdateParameters.Add("Sort", DbType.Int32, M_Entity_Update.Sort.ToString());
                this.ObjectDataSource1.UpdateParameters.Add("FY", DbType.String, M_Entity_Update.FY);
                this.ObjectDataSource1.UpdateParameters.Add("IsDel", DbType.Int32, M_Entity_Update.IsDel.ToString());
                this.ObjectDataSource1.UpdateParameters.Add("GXSJ", DbType.Date, M_Entity_Update.GXSJ.ToString());
                this.ObjectDataSource1.UpdateParameters.Add("ID", DbType.Int32, M_Entity_Update.ID.ToString());


                this.ObjectDataSource1.UpdateParameters.Add("BM", DbType.String, "");

                this.ObjectDataSource1.UpdateParameters.Add("SJ", DbType.String, "");


                this.ObjectDataSource1.Update();

                GridView1.EditIndex = e.RowIndex;
                //GridView1.EditIndex = -1;

                //记入日志
                TOOL.Login.OptObject = M_Entity_Update.ID.ToString();//操作ID
                TOOL.Login.BZ = "操作后:" + M_Entity_Update.Statistics();//备注

                OperationLog OperationLogAddupdate = new OperationLog();
                OperationLogAddupdate.OperationLogAdd("09", "04");
            }
            catch (Exception ex)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('" + ex.Message+ "');</script>");
            }
        }

        //取消
        protected void GridView1_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            GridView1.EditIndex = -1;

        }



        //流方式下载
        protected void btnDownload_Click(object sender, EventArgs e)
        {

            string sPath = AppDomain.CurrentDomain.BaseDirectory + "\\Excel\\部门管理.xls";

            bool b = File.Exists(sPath);
            if (!b)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('设备登记表格模版不存在，请联系管理员！');</script>");
                return;

            }


            string fileName = "部门管理.xls";//客户端保存的文件名


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

                //Model.Device M_Device = new Model.Device();

                Model.Entity M_Entity_Add = new Model.Entity();
                BLL.Entity B_Entity_Add = new BLL.Entity();
                StringBuilder strSql = new StringBuilder();
                string s_BMDM = "";

                for (int i = 0; i < dr.Length; i++)
                {

                    //M_Device.Manufacturer = dr[i]["厂家"].ToString();
                    M_Entity_Add.BMMC = dr[i]["单位名称"].ToString();
                    M_Entity_Add.BMQC = dr[i]["单位全称"].ToString();
                    M_Entity_Add.BMJC = dr[i]["单位简称"].ToString();

                    if (B_Entity_Add.ExistsBMDM(dr[i]["部门代码"].ToString()))
                    {
                        //strSql.Append(dr[i]["部门代码"].ToString());
                        //continue;
                        s_BMDM = dr[i]["部门代码"].ToString();//如果重复启动更新功能
                    
                    }


                    M_Entity_Add.BMDM = dr[i]["部门代码"].ToString();//不能重复

                    M_Entity_Add.SJBM = dr[i]["上级部门"].ToString();
                    M_Entity_Add.Depth = dr[i]["深度"].ToString();
                    M_Entity_Add.PicUrl = dr[i]["图片"].ToString();
                    M_Entity_Add.UserCount =int.Parse(dr[i]["用户数量"].ToString() == "" ? "0" : dr[i]["用户数量"].ToString());
                    M_Entity_Add.YZMC = dr[i]["印章名称"].ToString();
                    M_Entity_Add.FZJG = dr[i]["发证机关"].ToString();
                    M_Entity_Add.BMJB = dr[i]["部门级别"].ToString();



                    M_Entity_Add.LXDZ = dr[i]["联系地址"].ToString();
                    M_Entity_Add.FZR = dr[i]["负责人"].ToString();
                    M_Entity_Add.LXR = dr[i]["联系人"].ToString();
                    M_Entity_Add.LXDH = dr[i]["联系电话"].ToString();
                    M_Entity_Add.JKYH = dr[i]["缴款银行"].ToString();
                    M_Entity_Add.FY = dr[i]["法院"].ToString();
                    M_Entity_Add.FYJG = dr[i]["复议机构"].ToString();
                    M_Entity_Add.Lo =decimal.Parse( dr[i]["经度"].ToString() == "" ? "0.0" : dr[i]["经度"].ToString());
                    M_Entity_Add.La =decimal.Parse(dr[i]["纬度"].ToString() == "" ? "0.0" : dr[i]["纬度"].ToString());
                    M_Entity_Add.BZ = dr[i]["备注"].ToString();
            



                    try
                    {

                        M_Entity_Add.CJSJ = System.DateTime.Now;
                        int count = 0;
                        if (s_BMDM != "")
                        {
                            count = B_Entity_Add.UpdateEnter(M_Entity_Add);

                        }
                        else
                        {
                            count = B_Entity_Add.AddEnter(M_Entity_Add);
                        }
                        if (count > 0)
                        {
                            successly++;
                        }
                        else
                        {
                            strSql.Append(dr[i]["部门代码"].ToString());
                            continue;
                        
                        }

                    }
                    catch (Exception ex)
                    {
                        _Result = _Result + ex.Message + "\\n\\r";
                    }
                }


                if (strSql.Length > 0)
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些部门代码导入时出错:" + strSql.ToString() + ",');</script>");

                }


                if (successly == rowsnum)
                {

                    //记入日志
                    TOOL.Login.OptObject = "";//操作ID
                    TOOL.Login.BZ = "导入了" + successly + "条数据";//备注

                    OperationLog OperationLogAddupdate = new OperationLog();
                    OperationLogAddupdate.OperationLogAdd("09", "10");

                    string strmsg = "Excle表导入成功!";
                    System.Web.HttpContext.Current.Response.Write("<Script Language='JavaScript'>window.alert('" + strmsg + "');</script>");
                }
                else
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('Excle表导入失败!" + _Result + "');</script>");
                }
            }
        }









    }
}