using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Data.SqlClient;

namespace JingWuTong
{
    public partial class RemoteConfig : Form
    {
        public RemoteConfig()
        {
            InitializeComponent();
        }

        public string oradb="";
        private void btnTestLink_Click(object sender, EventArgs e)
        {
            try
            {
                string ip = txtIp.Text.Trim();
                string port = txtPort.Text.Trim();
                string database = txtDataBase.Text.Trim();
                string username = txtUserName.Text.Trim();
                string pwd = txtPwd.Text.Trim();

                oradb = "Data Source=(DESCRIPTION="
               + "(ADDRESS=(PROTOCOL=TCP)(HOST="+ip+")(PORT="+port+"))"
               + "(CONNECT_DATA=(SERVICE_NAME="+database+")));"
               + "User Id="+username+";Password="+pwd+";";
                OracleConnection conn = new OracleConnection(oradb);
                conn.Open(); 
                MessageBox.Show("连接成功");
                conn.Dispose();
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Program.oradb = oradb;
            MessageBox.Show("保存成功");
        }
    }
}
