using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using TOOL;

namespace JingWuTong
{
    public partial class login : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {

   

                //string strHostName = Dns.GetHostName(); //得到本机的主机名
                //IPHostEntry ipEntry = Dns.GetHostEntry(strHostName); //取得本机IP
                TOOL.Login.IP = GetServerIP().ToString();

                if (HttpContext.Current.Request.Cookies["cookieName"] != null)
                {

                    HttpCookie cookies = HttpContext.Current.Request.Cookies["cookieName"];
                    if (cookies["check"] != null && cookies["check"] == "1")
                    {

                        this.maintain.Checked = true;
                    }
                    if (cookies["JYBH"] != null && this.firstname.Value == "" && this.maintain.Checked == true)
                    {
                        this.firstname.Value = Server.UrlDecode(cookies["JYBH"]);
                    }
                    else
                    {
                        //HttpCookie cookies = new HttpCookie("cookieName");
                        cookies.Expires = DateTime.Now.AddDays(-1);
                        Response.Cookies.Add(cookies);
                    
                    }

                }


            }
        }


        ////获取IP4
        //public string GetLocalIP()
        //{
        //    try
        //    {
               
        //        string HostName = Dns.GetHostName(); //得到主机名
       

        //        IPHostEntry IpEntry = Dns.GetHostEntry(HostName);
                
        //        for (int i = 0; i < IpEntry.AddressList.Length; i++)
        //        {
        //            //从IP地址列表中筛选出IPv4类型的IP地址
        //            //AddressFamily.InterNetwork表示此IP为IPv4,
        //            //AddressFamily.InterNetworkV6表示此地址为IPv6类型
        //            if (IpEntry.AddressList[i].AddressFamily == AddressFamily.InterNetwork)
        //            {
        //                return IpEntry.AddressList[i].ToString();
        //            }
        //        }
        //        return "";
        //    }
        //    catch (Exception ex)
        //    {
        //        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('"+ex.Message+"');</script>");
        //        return "";
        //    }
        //}



        public  IPAddress GetServerIP()
        {
            IPAddress ipaddress = IPAddress.Parse("0.0.0.0");
            NetworkInterface[] interfaces = NetworkInterface.GetAllNetworkInterfaces();
         
                foreach (NetworkInterface ni in interfaces)
                {
                    if (ni.Name == "本地连接")
                    {
                        if (ni.NetworkInterfaceType == NetworkInterfaceType.Ethernet)
                        {
                            foreach (UnicastIPAddressInformation ip in ni.GetIPProperties().UnicastAddresses)
                            {
                                if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                                {
                                    ipaddress = ip.Address;
                                }
                            }
                        }
                    }
                }
            
            return ipaddress;
        }




        protected void Button2_Click(object sender, EventArgs e)
        {

            try
            {
                BLL.B_ACL_USER b_ACL_USER = new BLL.B_ACL_USER();

                string s_JYBH = this.firstname.Value;

                string s_Password = this.firstpassword.Value;



                if (b_ACL_USER.Exists(s_JYBH, s_Password) == null)
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('账号密码错误');</script>");

                    return;
                }

                int i_JSID = (int)b_ACL_USER.Exists(s_JYBH, s_Password);
                //ExistsXM

                if (i_JSID > 0)
                {
                    //添加Cookies

                    Model.M_ACL_USER model = new Model.M_ACL_USER();

                    model=b_ACL_USER.ExistsModel(s_JYBH, s_Password);



                        HttpCookie cookies = new HttpCookie("cookieName");
                        cookies["name"] = Server.UrlEncode(model.XM);
                        cookies["BMDM"] = Server.UrlEncode(model.BMDM);//部门编号
                        cookies["SJBM"] = Server.UrlEncode(model.SJBM);//部门级别

                        if (this.maintain.Checked == true)//保持登录状态
                        {
                            cookies["check"] = "1";
                        }
                        else {
                            cookies["check"] = "0";
                        
                        }

                        cookies["JYBH"] =Server.UrlEncode(s_JYBH);
                        TOOL.Login.OptObject =Server.UrlEncode(s_JYBH);
                        cookies.Expires = DateTime.Now.AddDays(1);
                        Response.Cookies.Add(cookies);




                    TOOL.Login.I_JSID = i_JSID;

                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('登录成功！');</script>");


                    try
                    {
                        //添加操作日志

                        OperationLog OperationLogAddupdate = new OperationLog();
                        OperationLogAddupdate.OperationLogAdd("00", "01");

                    }

                    catch (Exception ex)
                    {
                        Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('"+ex.Message+"');</script>");
                    
                    }




                    BLL.B_Role B_Role_Load = new BLL.B_Role();

                    string[] s_power = ((string)B_Role_Load.Exists(TOOL.Login.I_JSID)).Split('|');//权限参数




                    if (s_power[0] != "0")
                    {

                        //strjs.Append(" $('ul li:eq(0)').hide();"); //首页 
                        Response.Redirect("~/index.html");
                        return;

                    }


                    if (s_power[2] != "0")
                    {

                        // strjs.Append(" $('ul li:eq(1)').hide();"); //设备查看
                        Response.Redirect("~/Equipment/Equipment_Examine.aspx");
                        return;

                    }


                    if (s_power[4] != "0")
                    {

                        //strjs.Append(" $('ul li:eq(2)').hide();"); //实时状况
                        Response.Redirect("~/map.html");
                        return;

                    }

                    if (s_power[6] != "0")
                    {

                        //strjs.Append(" $('ul li:eq(3)').hide();"); //数据统计
                        Response.Redirect("~/dataManagement.html");
                        return;

                    }

                    if (s_power[10] != "0")
                    {

                        // strjs.Append(" $('ul li:eq(4)').hide();"); //设备管理
                        Response.Redirect("~/Equipment/Equipment_Management.aspx");
                        return;

                    }


                    if (s_power[18] != "0")
                    {

                        //strjs.Append(" $('ul li:eq(5)').hide();"); //人员管理
                        Response.Redirect("~/Personnel/Personnel_Manageme.aspx");
                        return;

                    }




                    if (s_power[24] != "0")
                    {
                        Response.Redirect("~/SystemSetup/SystemSetup.aspx");
                        return;

                        //strjs.Append(" $('ul li:eq(6)').hide();"); //系统设置

                    }




                    //Response.Redirect("~/Top.aspx");


                }
                else
                {
                    Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('请赋予账户权限');</script>");

                }
            }



            catch (Exception ex)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('"+ex.Message+"');</script>");
            
            }


        }

    }
}