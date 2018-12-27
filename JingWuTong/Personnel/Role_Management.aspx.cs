using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace JingWuTong.Personnel
{
    public partial class Role_Management : System.Web.UI.Page
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
                    this.GridView1.Columns[7].Visible = false;

                }

                if (s_power[21] == "0")//编辑
                {
                    this.GridView1.Columns[6].Visible = false;
                }

                if (s_power[22] == "0")//搜索
                {
                    this.Button1.Visible = false;
                }

                //if (s_power[23] == "0")//导入
                //{
                //    this.B_Into.Visible = false;
                //}

          



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

                        BLL.B_ACL_USER Find_B_ACL_USER = new BLL.B_ACL_USER();

                        int i_ID= int.Parse( e.CommandArgument.ToString());


                      bool b_JSID=Find_B_ACL_USER.ExistsJSID(i_ID);//查找是否绑定了警员



                      if (b_JSID == true)
                        {
                            Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('请将警员与角色解除绑定，在删除角色');</script>");
                            break;
                        }

                        else
                        {

                            BLL.B_Role B_Role_delet = new BLL.B_Role();
                            bool b_ID= B_Role_delet.Delete(i_ID);


                          if (b_ID == true)
                            {
                                //记入日志
                                TOOL.Login.OptObject = i_ID.ToString();//操作ID
                                TOOL.Login.BZ = "操作后:无";//备注

                                OperationLog OperationLogAddupdate = new OperationLog();
                                OperationLogAddupdate.OperationLogAdd("06", "05");

                                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('删除成功！');</script>");
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



        protected void Button2_Click(object sender, EventArgs e)
        {
            GridView1.PageIndex = 0;

        }



    }
}