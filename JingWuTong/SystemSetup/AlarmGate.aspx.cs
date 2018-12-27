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
    public partial class AlarmGate : System.Web.UI.Page
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
             


           
            }

        }





        protected void Button2_Click(object sender, EventArgs e)
        {

            this.Dr_DeviceNmae.Items[0].Value = "-2";
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


                Model.M_AlarmGate M_AlarmGate_Update = new Model.M_AlarmGate();



                M_AlarmGate_Update.CommonAlarmGate = int.Parse(((TextBox)(GridView1.Rows[e.RowIndex].Cells[4].FindControl("T_CommonAlarmGate"))).Text.ToString().Trim() == "" ? "0" : ((TextBox)(GridView1.Rows[e.RowIndex].Cells[4].FindControl("T_CommonAlarmGate"))).Text.ToString().Trim());

                M_AlarmGate_Update.UrgencyAlarmGate = int.Parse(((TextBox)(GridView1.Rows[e.RowIndex].Cells[5].FindControl("T_UrgencyAlarmGate"))).Text.ToString().Trim() == "" ? "0" : ((TextBox)(GridView1.Rows[e.RowIndex].Cells[5].FindControl("T_UrgencyAlarmGate"))).Text.ToString().Trim());


                M_AlarmGate_Update.ID = int.Parse(((Label)(GridView1.Rows[e.RowIndex].Cells[1].FindControl("L_ID"))).Text.ToString().Trim());



                //BLL.Device B_Device_update = new BLL.Device();
                //B_Device_update.UpdateRemind(M_Device_update);



                this.ObjectDataSource1.UpdateMethod = "Update";

                this.ObjectDataSource1.UpdateParameters.Add("CommonAlarmGate", DbType.Int32, M_AlarmGate_Update.CommonAlarmGate.ToString());
                this.ObjectDataSource1.UpdateParameters.Add("UrgencyAlarmGate", DbType.Int32, M_AlarmGate_Update.UrgencyAlarmGate.ToString());

                this.ObjectDataSource1.UpdateParameters.Add("ID", DbType.Int32, M_AlarmGate_Update.ID.ToString());


                this.ObjectDataSource1.UpdateParameters.Add("TypeName", DbType.String, "");

                this.ObjectDataSource1.UpdateParameters.Add("Name", DbType.String, "");


                this.ObjectDataSource1.Update();

                GridView1.EditIndex = e.RowIndex;
                //GridView1.EditIndex = -1;
            }
            catch (Exception ex)
            {

                Page.ClientScript.RegisterStartupScript(GetType(), "message", "<script>alert('" + ex.Message + "');</script>");
            }
        }

        //取消
        protected void GridView1_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            GridView1.EditIndex = -1;

        }
















    }
}