using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong.Personnel
{
    public partial class Police_Management : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {


                BLL.B_Role B_Role_Load = new BLL.B_Role();
                string[] s_power = ((string)B_Role_Load.Exists(TOOL.Login.I_JSID)).Split('|');//权限参数

                if (s_power[19] == "0")
                {
                    this.B_Add.Visible = false;  //新增

                }

                if (s_power[20] == "0") //删除
                {
                    this.GridView1.Columns[11].Visible = false;

                }

                if (s_power[21] == "0")//编辑
                {
                    this.GridView1.Columns[10].Visible = false;
                }

                if (s_power[22] == "0")//搜索
                {
                    this.Button1.Visible = false;
                }

                if (s_power[23] == "0")//导入
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


                //警员类型
                BLL.B_PoliceType dr_PoliceType = new BLL.B_PoliceType();

                this.Dr_JYLX.DataSource = dr_PoliceType.GetList();
                this.Dr_JYLX.DataTextField = "TypeName";
                this.Dr_JYLX.DataValueField = "ID";
                this.Dr_JYLX.DataBind();

                //this.Dr_JYLX.Items.Insert(0, new ListItem("---选择类型---", "-1"));
                this.Dr_JYLX.Items.Insert(0, new ListItem("全选", "-1"));
                this.Dr_JYLX.SelectedIndex = 0;

                //角色类型

                BLL.B_Role dr_Role = new BLL.B_Role();
                this.Dr_JSID.DataSource = dr_Role.GetList();
                this.Dr_JSID.DataTextField = "RoleName";
                this.Dr_JSID.DataValueField = "ID";
                this.Dr_JSID.DataBind();
                //this.Dr_JSID.Items.Insert(0, new ListItem("---选择类型---", "-1"));
                this.Dr_JSID.Items.Insert(0, new ListItem("全选", "-1"));
                this.Dr_JSID.SelectedIndex = 0;

                //警员职务

                BLL.Position dr_Position = new BLL.Position();

                this.Dr_LDJB.DataSource = dr_Position.GetList();
                this.Dr_LDJB.DataTextField = "PositionName";
                this.Dr_LDJB.DataValueField = "ID";
                this.Dr_LDJB.DataBind();
                //this.Dr_JSID.Items.Insert(0, new ListItem("---选择类型---", "-1"));
                this.Dr_LDJB.Items.Insert(0, new ListItem("全选", "-1"));
                this.Dr_LDJB.SelectedIndex = 0;

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

                        BLL.Device Find_Device = new BLL.Device();

                        BLL.B_ACL_USER Del_B_ACL_USE = new BLL.B_ACL_USER();


                        string[] estr = e.CommandArgument.ToString().Split(',');

                        long ParentID = long.Parse(estr[0]);
                        string S_JYBH = estr[1].ToString();

                        bool b_JYBH = Find_Device.ExistsJYBH(S_JYBH);//查找是否绑定设备

                        if (b_JYBH == true)
                        {
                            Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('请将警员与设备解除绑定，在删除警员');</script>");
                            break;
                        }

                        else
                        {

                            bool b_B_ACL_USE = Del_B_ACL_USE.Delete(ParentID);//删除警员表


                            if (b_B_ACL_USE == true)
                            {
                                //记入日志
                                TOOL.Login.OptObject = ParentID.ToString();//操作ID
                                TOOL.Login.BZ = "操作后:无";//备注

                                OperationLog OperationLogAddupdate = new OperationLog();
                                OperationLogAddupdate.OperationLogAdd("05", "05");

                                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('删除成功！'); </script>");
                                this.GridView1.DataBind();

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

        protected void Button2_Click(object sender, EventArgs e)
        {


            this.Dr_JYLX.Items[0].Value = "-2";
            this.Dr_JSID.Items[0].Value = "-2";
            this.Dr_LDJB.Items[0].Value = "-2";
            GridView1.PageIndex = 0;

        }





        //流方式下载
        protected void btnDownload_Click(object sender, EventArgs e)
        {

            string sPath = AppDomain.CurrentDomain.BaseDirectory + "\\Excel\\警员管理.xls";

            bool b = File.Exists(sPath);
            if (!b)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('设备登记表格模版不存在，请联系管理员！');</script>");
                return;

            }


            string fileName = "警员管理.xls";//客户端保存的文件名


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

                Model.M_ACL_USER M_ACL_USER_Add =new Model.M_ACL_USER();
                BLL.B_ACL_USER B_ACL_USER_Add = new BLL.B_ACL_USER();

                StringBuilder strSql = new StringBuilder();
                StringBuilder strSql2 = new StringBuilder();
                StringBuilder strSql3 = new StringBuilder();
                StringBuilder strSql4 = new StringBuilder();
                StringBuilder strSql5 = new StringBuilder();
                StringBuilder strSql6 = new StringBuilder();
                StringBuilder strSql7 = new StringBuilder();
                StringBuilder strSql8 = new StringBuilder();
                StringBuilder strSql9 = new StringBuilder();
                string s_jybh = "";

                for (int i = 0; i < dr.Length; i++)
                {
                    //前面除了你需要在建立一个“upfiles”的文件夹外，其他的都不用管了，你只需要通过下面的方式获取Excel的值，然后再将这些值用你的方式去插入到数据库里面

                    if (B_ACL_USER_Add.ExistsJYBH(dr[i]["警员编号"].ToString())) //不能重复
                    {
                        s_jybh = dr[i]["警员编号"].ToString();//如果重复启动更新功能

                        //strSql.Append(dr[i]["警员编号"].ToString());
                        //continue;

                    }
                   
                        M_ACL_USER_Add.JYBH = dr[i]["警员编号"].ToString();

                        if (string.IsNullOrEmpty(M_ACL_USER_Add.JYBH))
                        {
                            strSql.Append(dr[i]["警员编号"].ToString());
                            continue;
                        }




                    M_ACL_USER_Add.XM = dr[i]["姓名"].ToString();



                    BLL.Entity Entity_Find = new BLL.Entity();

                    object O_BMDM = Entity_Find.GetListBMMD(dr[i]["部门代码"].ToString());

                    if (O_BMDM == null)//如果输入的部门,名称有问题。
                    {
                        strSql2.Append(dr[i]["警员编号"].ToString());
                        continue;

                    }
                    else
                    {
                        M_ACL_USER_Add.BMDM = dr[i]["部门代码"].ToString();
                    }


                    //M_ACL_USER_Add.BMDM = dr[i]["部门名称"].ToString();

                    if (dr[i]["身份证号码"].ToString() != "")
                    {

                        if (!Regex.IsMatch(dr[i]["身份证号码"].ToString(), @"\d{17}[\d|X]|\d{15}", RegexOptions.ECMAScript))
                        {


                            strSql5.Append(dr[i]["警员编号"].ToString());
                            continue;

                        }
                    }

                    M_ACL_USER_Add.SFZMHM = dr[i]["身份证号码"].ToString();//正则

                    M_ACL_USER_Add.JTZZ = dr[i]["家庭地址"].ToString();

                    BLL.B_PoliceType B_B_PoliceType_Load = new BLL.B_PoliceType();

                  object O_JYLX=B_B_PoliceType_Load.ExistsJYLX(dr[i]["警员类型"].ToString());

                  if (O_JYLX== null)//如果输入的警员类型有问题。
                  {
                      strSql3.Append(dr[i]["警员编号"].ToString());
                      continue;

                  }
                  else
                  {
                      M_ACL_USER_Add.JYLX = (int)O_JYLX;
                  }


                  




                    M_ACL_USER_Add.MM = dr[i]["密码"].ToString();


                    BLL.B_Role B_Role_Load = new BLL.B_Role();

                     object O_JSID=B_Role_Load.ExistsJSID(dr[i]["角色"].ToString());

                     if (O_JSID == null)//如果输入的角色有问题。
                     {
                         strSql4.Append(dr[i]["警员编号"].ToString());
                         continue;

                     }
                     else
                     {
                         M_ACL_USER_Add.JSID =(int) O_JSID;

                     }


                     if (dr[i]["手机"].ToString() != "")
                     {

                         if (!Regex.IsMatch(dr[i]["手机"].ToString(), @"^1[3|4|5|7|8][0-9]{9}$", RegexOptions.ECMAScript))
                         {


                             strSql6.Append(dr[i]["警员编号"].ToString());
                             continue;
                         }
                     }

                    M_ACL_USER_Add.SJ= dr[i]["手机"].ToString(); //正则

                    if (dr[i]["办公电话"].ToString() != "")
                    {

                        if (!Regex.IsMatch(dr[i]["办公电话"].ToString(), @"(\(\d{3}\)|\d{3}-)?\d{8}", RegexOptions.ECMAScript))
                        {

                            strSql7.Append(dr[i]["警员编号"].ToString());
                            continue;
                        }

                    }

                    M_ACL_USER_Add.BGDH = dr[i]["办公电话"].ToString();//正则

                    M_ACL_USER_Add.BZLX =dr[i]["编制类型"].ToString();//1-公安编 2-事业编 3-地方编

                    M_ACL_USER_Add.YWGW = dr[i]["业务岗位"].ToString();// 01-城区管理执勤 02-国省道执勤 03-高速执勤 04-县乡执勤 05-事故处理 06-车驾管 07-道路宣传 08-科技管理 09-其他

                    M_ACL_USER_Add.SGCLDJ=dr[i]["事故处理等级"].ToString();//0:无等级,1:高级,2:中级,3:初级;

                    BLL.Position B_Position_load = new BLL.Position();
                    object O_LDJB = B_Position_load.GetListLDJB(dr[i]["领导级别"].ToString());
                    if (O_LDJB == null)//如果输入的领导级别有问题
                    {
                        strSql8.Append(dr[i]["警员编号"].ToString());
                        continue;

                    }
                    else
                    {

                        M_ACL_USER_Add.LDJB = O_LDJB.ToString();//D0,总队领导D1,支队领导D2,大队领导D3,中队领导ZZ,其他
                    }

                    if (dr[i]["出生日期"].ToString() == "")
                    {
                        M_ACL_USER_Add.CSRQ = null;

                    }
                    else
                    {
                        M_ACL_USER_Add.CSRQ = Convert.ToDateTime(dr[i]["出生日期"].ToString());
                    }



                    M_ACL_USER_Add.XB = dr[i]["性别"].ToString();
                    M_ACL_USER_Add.JX= dr[i]["警衔"].ToString();

                    M_ACL_USER_Add.JB = dr[i]["级别"].ToString();


                    M_ACL_USER_Add.JG = dr[i]["籍贯"].ToString();

                    if (dr[i]["入党团时间"].ToString() == "")
                    {

                        M_ACL_USER_Add.RDTSJ =null;

                    }

                    else
                    {
                        M_ACL_USER_Add.RDTSJ = Convert.ToDateTime(dr[i]["入党团时间"].ToString());
                    
                    }
                    M_ACL_USER_Add.ZZMM = dr[i]["政治面貌"].ToString();


                    M_ACL_USER_Add.MZ = dr[i]["民族"].ToString();
                    M_ACL_USER_Add.XL = dr[i]["学历"].ToString();

                    M_ACL_USER_Add.ZY = dr[i]["专业"].ToString();
                    M_ACL_USER_Add.ZW = dr[i]["岗位职务级别"].ToString();

                    if (dr[i]["入队时间"].ToString() == "")
                    {
                        M_ACL_USER_Add.RDTSJ = null;

                    }
                    else
                    {
                        M_ACL_USER_Add.RDTSJ = Convert.ToDateTime(dr[i]["入队时间"].ToString());
                    }

                    if (dr[i]["参加工作时间"].ToString() == "")
                    {
                        M_ACL_USER_Add.CGSJ = null;
                    }
                    else
                    {
                        M_ACL_USER_Add.CGSJ = Convert.ToDateTime(dr[i]["参加工作时间"].ToString());
                    }

                    if (dr[i]["任现职时间"].ToString() == "")
                    {
                        M_ACL_USER_Add.RXZSJ = null;
                    }
                    else
                    {
                        M_ACL_USER_Add.RXZSJ = Convert.ToDateTime(dr[i]["任现职时间"].ToString());
                    }

                    M_ACL_USER_Add.ZFZGDJ = dr[i]["执法资格等级"].ToString();

                    M_ACL_USER_Add.ZT = int.Parse(dr[i]["状态（有效或无效）"].ToString() == "" ? "1" : dr[i]["状态（有效或无效）"].ToString());


                    M_ACL_USER_Add.CJSJ = System.DateTime.Now;


                    try
                    {
                        int count = 0;
                        if (s_jybh != "")
                        {
                            count = B_ACL_USER_Add.UpdateData(M_ACL_USER_Add);

                        }
                        else
                        {
                            count = B_ACL_USER_Add.Add(M_ACL_USER_Add);
                        }
                        if (count > 0)
                        {
                            successly++;
                        }
                        else
                        {

                            strSql9.Append(dr[i]["警员编号"].ToString());
                            continue;
                        }

                    }
                    catch (Exception ex)
                    {
                        _Result = _Result + ex.InnerException + "\\n\\r";
                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('" + _Result + "');</script>");
                    }
                }

                if (strSql.Length > 0)
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些警员编号为空:" + strSql.ToString() + ",');</script>");


                }


                if (strSql2.Length > 0)
                {


                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些警员编号输入的部门代码无法查询到:" + strSql2.ToString() + ",');</script>");
                }


                if (strSql3.Length > 0)
                {


                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些警员编号输入的警员类型无法查询到:" + strSql3.ToString() + ",');</script>");
                }


                if (strSql4.Length > 0)
                {


                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些警员编号输入的角色无法查询到:" + strSql4.ToString() + ",');</script>");
                }


                if (strSql5.Length > 0)
                {


                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些警员编号对应的身份证号码格式不正确:" + strSql5.ToString() + ",');</script>");
                }


                if (strSql6.Length > 0)
                {


                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些警员编号对应的手机号码格式不正确:" + strSql6.ToString() + ",');</script>");
                }


                if (strSql7.Length > 0)
                {


                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些警员编号对应的办公室电话格式不正确:" + strSql7.ToString() + ",');</script>");
                }

                if (strSql8.Length > 0)
                {


                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些警员编号对应的领导级别不正确:" + strSql8.ToString() + ",');</script>");
                }
                if (strSql9.Length > 0)
                {


                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('以下这些警员编号导入时出错:" + strSql9.ToString() + ",');</script>");
                }


                if (successly == rowsnum)
                {

                    //记入日志
                    TOOL.Login.OptObject = "";//操作ID
                    TOOL.Login.BZ = "导入了" + successly + "条数据";//备注
                    OperationLog OperationLogAddupdate = new OperationLog();
                    OperationLogAddupdate.OperationLogAdd("05", "10");

                    string strmsg = "Excle表导入成功!";


                    System.Web.HttpContext.Current.Response.Write("<Script Language='JavaScript'>window.alert('" + strmsg + "');</script>");
                }
                else
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('Excle表导入失败!');</script>");
                }
            }
        }





    }
}