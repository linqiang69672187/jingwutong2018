using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Threading;

namespace JingWuTong
{
    public partial class Form1 : Form
    {
        //public string oradb = "Data Source=(DESCRIPTION="
        //     + "(ADDRESS=(PROTOCOL=TCP)(HOST=10.8.59.253)(PORT=1521))"
        //     + "(CONNECT_DATA=(SERVICE_NAME=orcl)));"
        //     + "User Id=SYSTEM;Password=51411482;";

        public string oradb = "";
        public SqlConnection conn = new SqlConnection(Program.sqladb);
        public Int64 historyGpsRevice = 0;
        public Int64 GpsRevice = 0;
        public Int64 violationRevice = 0;
        public Int64 forceRevice = 0;
        public Int64 devRevice = 0;
        public Int64 duiJiangJiRecive = 0;
        public Int64 duiJiangJiRecive1 = 0;
        public List<GPSData> gpsList = new List<GPSData>();
        private UdpClient ReceiveClient;
        //private Thread ReceiveThread;
        //private ThreadState threadStatus;
        private Mutex mut = new Mutex();
        SqlBulkCopy bulkCopy;
        //public ThreadState ThreadStatus
        //{
        //    get { return threadStatus; }
        //}

        public Form1()
        {
            InitializeComponent();
            this.textBox1.Text = "准备就绪";
            WriteLog("警务通接收程序启动");
            //threadStatus = ThreadState.Unstarted;
            //ReceiveThread = new Thread(new ThreadStart(DuiJiangJiDataHandler));

           bulkCopy = new SqlBulkCopy(Program.sqladb, SqlBulkCopyOptions.KeepIdentity | SqlBulkCopyOptions.FireTriggers);
          //  bulkCopy = new SqlBulkCopy(Program.sqladb, SqlBulkCopyOptions.KeepIdentity );
            bulkCopy.DestinationTableName = "HistoryGps";

            conn.Open();
        }


        private void btnRecive_Click(object sender, EventArgs e)
        {
            try
            {
                oradb = Program.oradb;
             //   if (oradb != "")
                if(true)
                {
                   // OracleConnection conn = new OracleConnection(oradb);
                  //  conn.Open();
                   // conn.Dispose();
                    gpsTimer.Start();
                    showTimer.Start();
                    //异步委托
                    DelegateProgress wtGps = new DelegateProgress(gpsDataHandler);
                    IAsyncResult syncGps = wtGps.BeginInvoke(null, null);
                    DelegateProgress wtVio = new DelegateProgress(vioDataHandler);
                    IAsyncResult syncVio = wtVio.BeginInvoke(null, null);
                    DelegateProgress wtDev = new DelegateProgress(devDataHandler);
                    IAsyncResult syncDev = wtDev.BeginInvoke(null, null);

                    DelegateProgress wtDuiJiangJi = new DelegateProgress(DuiJiangJiDataHandler);
                    IAsyncResult syncDuiJiangJi = wtDuiJiangJi.BeginInvoke(null, null);

                    //队列插入
                    DelegateProgress wtQueue = new DelegateProgress(QueueInsert);
                    IAsyncResult synQueue = wtQueue.BeginInvoke(null, null);
                    //if (threadStatus == ThreadState.Unstarted || threadStatus == ThreadState.Stopped)
                    //{
                    //    ReceiveThread.Start();
                    //    threadStatus = ThreadState.Running;
                    //}

                }
                else
                {
                    WriteLog("oracle数据库连接未设置");
                    MessageBox.Show("请先进行远程数据库的设置。");

                }
            }
            catch (Exception ex)
            {
                //WriteLog(ex.Message);
                MessageBox.Show("远程数据库连接不成功。");
            }
        }
        public delegate void DelegateProgress();
        public void gpsDataHandler()
        {
            int xh = 1;
            while (xh == 1)
            {
                return;
                try
                {

                    OracleConnection conn = new OracleConnection(oradb);
                    conn.Open();
                    //SqlConnection sqlCon = new SqlConnection(Program.sqladb);
                    //sqlCon.Open();
                    DateTime currentTime = DateTime.Now;
                    string curTime = currentTime.ToString("yyyy-MM-dd");
                    //调试
                    //string sql = "select * from SYS_HEARTBEAT where QQSJ>sysdate-10/1440 order by QQSJ DESC";
                    //string sql = "select * from SYS_HEARTBEAT";
                    //发布
                    string sql = "select * from tzydjw.v_to3rd_gps where QQSJ>sysdate-10/1440 order by QQSJ DESC";
                    OracleCommand cmd = new OracleCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;

                    //大容量数据，可以进行修改，使用DataSet和DataAdapter
                    DataSet ds = new DataSet();
                    OracleDataAdapter da = new OracleDataAdapter();
                    da.SelectCommand = cmd;
                    da.Fill(ds);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        string lon = "";
                        string lat = "";
                        string strGps = ds.Tables[0].Rows[i]["GPS"].ToString();
                        if (strGps.IndexOf(",") > 0)
                        {
                            lon = strGps.Substring(0, strGps.IndexOf(",")).Trim();
                            lat = strGps.Substring(strGps.IndexOf(",") + 1).Trim();

                        }
                        string padid = ds.Tables[0].Rows[i]["PDAID"].ToString();
                        string QQSJ = ds.Tables[0].Rows[i]["QQSJ"].ToString();

                        int deviceType = 0;
                        string contacts = "";
                        string devSql = string.Format("select DevType,Contacts from Device where DevId='{0}'", padid);
                        System.Data.DataTable dtDev = ExecuteRead(CommandType.Text, devSql, "Dev");
                        if (dtDev.Rows.Count > 0)
                        {
                            deviceType = int.Parse(dtDev.Rows[0][0].ToString());
                            contacts = dtDev.Rows[0][1].ToString();
                        }

                        string selSql = string.Format("select * from [JingWuTong].[dbo].[HistoryGps] where PDAID='{0}' and QQSJ='{1}'", padid, QQSJ);
                        System.Data.DataTable dt =ExecuteRead(CommandType.Text, selSql, "aa");
                        if (dt.Rows.Count == 0)
                        {
                        //mut.WaitOne();
                        string intsetsql = "Insert into HistoryGps([PDAID],[YHDM],[QQSJ],[YDSJ],[PDATIME],[SERVERTIME],[PDAVERSION],[SFGX],[SFQZGX],[YL1],[YL2],[YL3],[YL4],[Lo],[La],[DevType],[Contacts]) values(" +
                                  " '" + ds.Tables[0].Rows[i]["PDAID"].ToString() + "', '" + ds.Tables[0].Rows[i]["YHDM"].ToString() + "', '" + ds.Tables[0].Rows[i]["QQSJ"].ToString() +
                                  "', '" + ds.Tables[0].Rows[i]["YDSJ"].ToString() + "', '" + ds.Tables[0].Rows[i]["PDATIME"].ToString() + "', '" + ds.Tables[0].Rows[i]["SERVERTIME"].ToString() + "', '" + ds.Tables[0].Rows[i]["PDAVERSION"].ToString() +
                                  "', '" + ds.Tables[0].Rows[i]["SFGX"].ToString() + "', '" + ds.Tables[0].Rows[i]["SFQZGX"].ToString() + "', '" + ds.Tables[0].Rows[i]["YL1"].ToString() + "', '" + ds.Tables[0].Rows[i]["YL2"].ToString() +
                                  "', '" + ds.Tables[0].Rows[i]["YL3"].ToString() + "', '" + ds.Tables[0].Rows[i]["YL4"].ToString() + "', '" + lon + "','" + lat + "','4','" + contacts + "')";

                        ExecuteNonQuery(intsetsql);
                        //mut.ReleaseMutex();
                        historyGpsRevice++;

                        }
                        System.Threading.Thread.Sleep(50);


                    }


                    conn.Dispose(); //Close()也可以。
                    //sqlCon.Dispose();

                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.StackTrace);
                    WriteLog(ex.StackTrace);
                }
            }
        }
        private void gpsTimer_Tick(object sender, EventArgs e)
        {


        }

        public void vioDataHandler()
        {
            return;
            int xh = 1;
            while (xh == 1)
            {
                try
                {
                    OracleConnection conn = new OracleConnection(oradb);
                    conn.Open();
                    SqlConnection sqlCon = new SqlConnection(Program.sqladb);
                    sqlCon.Open();

                    //接收VIO_VIOLATION_IN数据
                    DateTime currentTime = DateTime.Now;
                    string curTime = currentTime.ToString("yyyy-MM-dd");
                    //调试
                    //string sql = "select * from VIO_VIOLATION_IN where LRSJ>sysdate-10/1440";
                    //string sql = "select * from VIO_VIOLATION_IN";
                    //发布
                    string sql = "select * from tzydjw.v_to3rd_vio_violation where LRSJ>sysdate-10/1440";
                    OracleCommand cmd = new OracleCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;

                    //大容量数据，可以进行修改，使用DataSet和DataAdapter
                    DataSet ds = new DataSet();
                    OracleDataAdapter da = new OracleDataAdapter();
                    da.SelectCommand = cmd;
                    da.Fill(ds);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        string lon = ds.Tables[0].Rows[i]["JD"].ToString();
                        string lat = ds.Tables[0].Rows[i]["WD"].ToString();

                        string padid = ds.Tables[0].Rows[i]["PDAID"].ToString();
                        string LRSJ = ds.Tables[0].Rows[i]["LRSJ"].ToString();

                        int deviceType = 0;
                        string devSql = string.Format("select DevType from Device where DevId='{0}'", padid);
                        System.Data.DataTable dtDev = ExecuteRead(CommandType.Text, devSql, "Dev");
                        if (dtDev.Rows.Count > 0)
                        {
                            deviceType = int.Parse(dtDev.Rows[0][0].ToString());
                        }

                        string selSql = string.Format("select * from [JingWuTong].[dbo].[VIO_VIOLATION_IN] where PDAID='{0}' and LRSJ='{1}'", padid, LRSJ);
                        System.Data.DataTable dt = ExecuteRead(CommandType.Text, selSql, "aa");
                        if (dt.Rows.Count == 0)
                        {
                            string intsetsql = "Insert into VIO_VIOLATION_IN([ZT],[LRSJ],[PDAMS],[JDSBH],[JSZH],[DSR],[HPZL],[HPHM],[WFSJ],[WFDZ],[WFXW],[ZQMJ],[FXJG],[Lo],[La],[PDAID],[DevType]) values(" +
                                      " '" + ds.Tables[0].Rows[i]["ZT"].ToString() + "', '" + ds.Tables[0].Rows[i]["LRSJ"].ToString() + "', '" + ds.Tables[0].Rows[i]["PDAMS"].ToString() +
                                      "', '" + ds.Tables[0].Rows[i]["JDSBH"].ToString() + "', '" + ds.Tables[0].Rows[i]["JSZH"].ToString() + "', '" + ds.Tables[0].Rows[i]["DSR"].ToString() + "', '" + ds.Tables[0].Rows[i]["HPZL"].ToString() +
                                      "', '" + ds.Tables[0].Rows[i]["HPHM"].ToString() + "', '" + ds.Tables[0].Rows[i]["WFSJ"].ToString() + "', '" + ds.Tables[0].Rows[i]["WFDZ"].ToString() + "', '" + ds.Tables[0].Rows[i]["WFXW"].ToString() +
                                      "', '" + ds.Tables[0].Rows[i]["ZQMJ"].ToString() + "', '" + ds.Tables[0].Rows[i]["FXJG"].ToString() + "', '" + lon + "','" + lat + "','" + padid + "','" + deviceType + "')";

                            ExecuteNonQuery(intsetsql);
                            violationRevice++;
                        }

                        System.Threading.Thread.Sleep(25);

                    }


                    //接收VIO_FORCE_IN数据
                    //调试
                    //sql = "select * from VIO_FORCE_IN where LRSJ>sysdate-10/1440";
                    //sql = "select * from VIO_FORCE_IN";
                    //发布
                    sql = "select * from tzydjw.v_to3rd_vio_force where LRSJ>sysdate-10/1440";
                    cmd = new OracleCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;

                    //大容量数据，可以进行修改，使用DataSet和DataAdapter
                    ds = new DataSet();
                    da = new OracleDataAdapter();
                    da.SelectCommand = cmd;
                    da.Fill(ds);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        string lon = ds.Tables[0].Rows[i]["JD"].ToString();
                        string lat = ds.Tables[0].Rows[i]["WD"].ToString();

                        string padid = ds.Tables[0].Rows[i]["PDAID"].ToString();
                        string LRSJ = ds.Tables[0].Rows[i]["LRSJ"].ToString();
                        int deviceType = 0;
                        string devSql = string.Format("select DevType from Device where DevId='{0}'", padid);
                        System.Data.DataTable dtDev = ExecuteRead(CommandType.Text, devSql, "Dev");
                        if (dtDev.Rows.Count > 0)
                        {
                            deviceType = int.Parse(dtDev.Rows[0][0].ToString());
                        }
                        string selSql = string.Format("select * from [JingWuTong].[dbo].[VIO_FORCE_IN] where PDAID='{0}' and LRSJ='{1}'", padid, LRSJ);
                        System.Data.DataTable dt = ExecuteRead(CommandType.Text, selSql, "aa");
                        if (dt.Rows.Count == 0)
                        {
                            string intsetsql = "Insert into VIO_FORCE_IN([ZT],[LRSJ],[PZBH],[JSZH],[DSR],[HPHM],[HPZL],[WFSJ],[WFDZ],[WFXW1],[SCZ1],[BZZ1],[WFXW2],[SCZ2],[BZZ2],[WFXW3],[SCZ3],[BZZ3],[WFXW4],[SCZ4],[BZZ4],[WFXW5],[SCZ5],[BZZ5],[FXJG],[ZQMJ],[Lo],[La],[PDAID],[DevType]) values(" +
                                      " '" + ds.Tables[0].Rows[i]["ZT"].ToString() + "', '" + ds.Tables[0].Rows[i]["LRSJ"].ToString() + "', '" + ds.Tables[0].Rows[i]["PZBH"].ToString() +
                                      "', '" + ds.Tables[0].Rows[i]["JSZH"].ToString() + "', '" + ds.Tables[0].Rows[i]["DSR"].ToString() + "', '" + ds.Tables[0].Rows[i]["HPHM"].ToString() + "', '" + ds.Tables[0].Rows[i]["HPZL"].ToString() +
                                      "', '" + ds.Tables[0].Rows[i]["WFSJ"].ToString() + "', '" + ds.Tables[0].Rows[i]["WFDZ"].ToString() + "', '" + ds.Tables[0].Rows[i]["WFXW1"].ToString() +
                                      "', '" + ds.Tables[0].Rows[i]["SCZ1"].ToString() + "', '" + ds.Tables[0].Rows[i]["BZZ1"].ToString() + "', '" + ds.Tables[0].Rows[i]["WFXW2"].ToString() +
                                      "', '" + ds.Tables[0].Rows[i]["SCZ2"].ToString() + "', '" + ds.Tables[0].Rows[i]["BZZ2"].ToString() + "', '" + ds.Tables[0].Rows[i]["WFXW3"].ToString() +
                                      "', '" + ds.Tables[0].Rows[i]["SCZ3"].ToString() + "', '" + ds.Tables[0].Rows[i]["BZZ3"].ToString() + "', '" + ds.Tables[0].Rows[i]["WFXW4"].ToString() +
                                      "', '" + ds.Tables[0].Rows[i]["SCZ4"].ToString() + "', '" + ds.Tables[0].Rows[i]["BZZ4"].ToString() + "', '" + ds.Tables[0].Rows[i]["WFXW5"].ToString() +
                                      "', '" + ds.Tables[0].Rows[i]["SCZ5"].ToString() + "', '" + ds.Tables[0].Rows[i]["BZZ5"].ToString() + "', '" + ds.Tables[0].Rows[i]["FXJG"].ToString() +
                                      "', '" + ds.Tables[0].Rows[i]["ZQMJ"].ToString() + "', '" + lon + "','" + lat + "','" + padid + "','" + deviceType + "')";

                            ExecuteNonQuery(intsetsql);
                            forceRevice++;
                        }
                        System.Threading.Thread.Sleep(25);
                    }

                    conn.Dispose(); //Close()也可以。
                    sqlCon.Dispose();
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.StackTrace);
                    WriteLog(ex.StackTrace);
                }
            }
        }

        public void devDataHandler()
        {
            return;
            int xh = 1;
            while (xh == 1)
            {
                try
                {
                    devRevice = 0;
                    OracleConnection conn = new OracleConnection(oradb);
                    conn.Open();
                    SqlConnection sqlCon = new SqlConnection(Program.sqladb);
                    sqlCon.Open();
                    //调试
                    //string sql = "select * from PDA_DEV order by GXSJ DESC";
                    //发布
                    string sql = "select * from tzydjw.v_to3rd_PDA_DEV order by GXSJ DESC";
                    OracleCommand cmd = new OracleCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;

                    //大容量数据，可以进行修改，使用DataSet和DataAdapter
                    DataSet ds = new DataSet();
                    OracleDataAdapter da = new OracleDataAdapter();
                    da.SelectCommand = cmd;
                    da.Fill(ds);
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        string simid = ds.Tables[0].Rows[i]["SIMID"].ToString();
                        //string yhdm = ds.Tables[0].Rows[i]["ZRR"].ToString();
                        ////获取单位全称,联系人
                        ////调试
                        ////sql = string.Format("SELECT d.BMMC,u.XM from acl_user u inner join acl_dept d on u.BMDM=d.BMDM where u.YHDM='{0}'", yhdm);
                        ////发布
                        //sql = string.Format("SELECT d.BMMC,u.XM from tzydjw.v_to3rd_acl_user u left join tzydjw.v_to3rd_acl_dept d on u.BMDM=d.BMDM where u.YHDM='{0}'", yhdm);
                        //cmd = new OracleCommand(sql, conn);
                        //cmd.CommandType = CommandType.Text;
                        //DataSet dsEntity = new DataSet();
                        //OracleDataAdapter daEntity = new OracleDataAdapter();
                        //daEntity.SelectCommand = cmd;
                        //daEntity.Fill(dsEntity);
                        //string strEntity="";
                        //string contacts="";
                        //if (dsEntity.Tables[0].Rows.Count > 0)
                        //{
                        //    strEntity = dsEntity.Tables[0].Rows[0][0].ToString().Trim();
                        //    contacts = dsEntity.Tables[0].Rows[0][1].ToString();
                        //}

                        if (ds.Tables[0].Rows[i]["PDAID"].ToString() == "")
                            continue;

                        string devSql = string.Format("select * from Device_Jwt where SIMID='{0}'", simid);
                        System.Data.DataTable dtDev = ExecuteRead(CommandType.Text, devSql, "DevJwt");
                        if (dtDev.Rows.Count > 0)
                        {
                            string updateSql = "update Device_JWT set Tel='" + ds.Tables[0].Rows[i]["SJHM"].ToString() + "',ZT=" + ds.Tables[0].Rows[i]["ZT"].ToString() + ",IMEI='"
                                + ds.Tables[0].Rows[i]["IMEI"].ToString() + "',DevId='" + ds.Tables[0].Rows[i]["PDAID"].ToString() + "' where SIMID='" + simid + "'";
                            ExecuteNonQuery(updateSql);
                        }
                        else
                        {

                            string existSql = string.Format("select * from Device_JWT where DevId='{0}'", ds.Tables[0].Rows[i]["PDAID"].ToString());
                            System.Data.DataTable existDev = ExecuteRead(CommandType.Text, existSql, "DevJwt1");
                            if (existDev.Rows.Count > 0)
                            {
                                sql = string.Format("delete from Device_JWT where DevId='{0}'", ds.Tables[0].Rows[i]["PDAID"].ToString());
                                ExecuteNonQuery(sql);
                            }

                            string intsetsql = "insert into Device_JWT([DevId],[DevType],[Tel],[ZT],[IMEI],[SIMID])values('" + ds.Tables[0].Rows[i]["PDAID"].ToString() + "',4,'"
                                + ds.Tables[0].Rows[i]["SJHM"].ToString() + "','" + ds.Tables[0].Rows[i]["ZT"].ToString() + "','" + ds.Tables[0].Rows[i]["IMEI"].ToString() +
                                "','" + ds.Tables[0].Rows[i]["SIMID"].ToString() + "')";
                            //WriteLog(intsetsql);
                            ExecuteNonQuery(intsetsql);
                            devRevice++;

                        }

                        System.Threading.Thread.Sleep(1000);

                    }

                    string deviceCyleSql = "select cycle from Device_JwtCycle";
                    int cycle = int.Parse(ExecuteScalar(deviceCyleSql).ToString());
                    System.Threading.Thread.Sleep(cycle * 60 * 1000);
                }
                catch (Exception ex)
                {
                    WriteLog(ex.StackTrace);
                }
            }
        }

        #region 原来的DuiJiangJiDataHandler
        public void DuiJiangJiDataHandler()
        {
            int xh = 1;
            if (ReceiveClient != null)
            {
                ReceiveClient.Close();
            }

            IPEndPoint remote = new IPEndPoint(IPAddress.Any, 0);
            ReceiveClient = new UdpClient(6800);
            duiJiangJiRecive = 0;
            duiJiangJiRecive1 = 0;
            while (xh == 1)
            {
                try
                {
                    //WriteLog("写入对讲机");
                    //Console.WriteLine("写入对讲机.");
                    /**
                                      GPSData gps = new GPSData();
                                      gps.PDAID = "331022000020670";
                                      gps.QQSJ = DateTime.Parse("2017-07-05 00:00:01.000");
                                      gps.YDSJ = DateTime.Parse("1900-01-01 00:00:00.000");
                                      gps.LA = 121.36273;
                                      gps.LO = 28.09808;
                                      gpsList.Add(gps);

                                      //WriteLog("issi:" + gps.PDAID + ";qqsj:" + gps.QQSJ.ToString() + ";YDSJ:" + gps.YDSJ.ToString() + ";LAT:" + gps.LA.ToString() + ";Lon:" + gps.LO.ToString());

                                      //string selSql = string.Format("select * from [JingWuTong].[dbo].[HistoryGps] where PDAID='{0}' and QQSJ='{1}'", issi, dd);
                                      //System.Data.DataTable dt = SQLHelper.ExecuteRead(CommandType.Text, selSql, "aa");
                                      //if (dt.Rows.Count == 0)
                                      //{
                                      //    mut.WaitOne();
                                      //    string intsetsql = "Insert into HistoryGps([PDAID],[QQSJ],[YDSJ],[Lo],[La],[DevType]) values(" +
                                      //              " '" + issi + "', '" + dd.ToString() + "', '1900-01-01 00:00:00.000', '" + lat + "','" + lon + "','2')";

                                      //    SQLHelper.ExecuteNonQuery(intsetsql); 
                                      //    duiJiangJiRecive++;
                                      //    mut.ReleaseMutex();

                                      //}
                                      duiJiangJiRecive1++;
                                      System.Threading.Thread.Sleep(10);
                      **/


                                   byte[] bytes = null;
                                      // GPSData data;

                                      bytes = ReceiveClient.Receive(ref remote);

                                      //WriteLog("bytes vales:" + ByteToString(bytes));

                                      //string strBytes = "AA AA CC CC 03 00 35 00 00 00 35 37 36 31 31 32 30 31 00 00 00 00 00 00 00 00 00 00 00 00 18 44 A4 A6 5D 5C 5E 40 E1 25 38 F5 81 A8 3C 40 00 00 20 00 00 00 01 00 E1 07 08 04 0B 04 15 00 00";

                                      //string strBytes = "aa aa cc cc 03 00 35 00 00 00 35 37 36 31 33 31 31 39 00 00 00 00 00 00 00 00 00 00 00 00 b6 80 d0 7a f8 56 5e 40 02 81 ce a4 4d 95 3c 40 00 00 43 00 00 00 01 00 e1 07 08 08 0e 1e 09 00 00";

                                      //strBytes = strBytes.Replace(" ", "");
                                      //bytes = DuiJiangJi.HexStrTobyte(strBytes);

                                      string issi = System.Text.Encoding.ASCII.GetString(bytes, 10, 20);
                                      issi = issi.Replace("\0", "");
                                      //经度
                                      double lon = 0.00000;
                                      Int32 size = Marshal.SizeOf(typeof(double));
                                      IntPtr buffer = Marshal.AllocHGlobal(size);
                                      try
                                      {
                                          Marshal.Copy(DuiJiangJi.getBbytes(bytes, 30, 8), 0, buffer, size);
                                          lon = (double)Marshal.PtrToStructure(buffer, typeof(double));
                                      }
                                      catch (Exception ex)
                                      {
                                          WriteLog(ex.Message);
                                      }
                                      finally
                                      {
                                          Marshal.FreeHGlobal(buffer);
                                          //GC.Collect();
                                      }

                                      //纬度
                                      double lat = 0.00000;
                                      size = Marshal.SizeOf(typeof(double));
                                      buffer = Marshal.AllocHGlobal(size);
                                      try
                                      {
                                          Marshal.Copy(DuiJiangJi.getBbytes(bytes, 38, 8), 0, buffer, size);
                                          lat = (double)Marshal.PtrToStructure(buffer, typeof(double));
                                      }
                                      catch (Exception ex)
                                      {
                                          WriteLog(ex.Message);
                                      }
                                      finally
                                      {
                                          Marshal.FreeHGlobal(buffer);
                                          //GC.Collect();
                                      }

                                      //年
                                      ushort year = 1900;
                                      size = Marshal.SizeOf(typeof(ushort));
                                      buffer = Marshal.AllocHGlobal(size);
                                      try
                                      {
                                          Marshal.Copy(DuiJiangJi.getBbytes(bytes, 54, 2), 0, buffer, size);
                                          year = (ushort)Marshal.PtrToStructure(buffer, typeof(ushort));
                                      }
                                      catch (Exception ex)
                                      {
                                          WriteLog(ex.Message);
                                      }
                                      finally
                                      {
                                          Marshal.FreeHGlobal(buffer);
                                          //GC.Collect();
                                      }

                                      string time = year + "-" + DuiJiangJi.getBbytes(bytes, 56, 1)[0].ToString() + "-" + DuiJiangJi.getBbytes(bytes, 57, 1)[0].ToString() +
                                  " " + DuiJiangJi.getBbytes(bytes, 58, 1)[0].ToString() + ":" + DuiJiangJi.getBbytes(bytes, 59, 1)[0].ToString() + ":" + DuiJiangJi.getBbytes(bytes, 60, 1)[0].ToString();

                                      DateTime dd = DateTime.Parse(time);
                                      duiJiangJiRecive1++;

                                      GPSData gps = new GPSData();
                                      gps.PDAID = issi;
                                      gps.QQSJ = dd;
                                      gps.YDSJ = DateTime.Parse("1900-01-01 00:00:00.000");
                                      gps.LA = lat;
                                      gps.LO = lon;
                                      gpsList.Add(gps);

                                      //WriteLog("issi:" + gps.PDAID + ";qqsj:" + gps.QQSJ.ToString() + ";YDSJ:" + gps.YDSJ.ToString() + ";LAT:" + gps.LA.ToString() + ";Lon:" + gps.LO.ToString());
                    
                                      //string selSql = string.Format("select * from [JingWuTong].[dbo].[HistoryGps] where PDAID='{0}' and QQSJ='{1}'", issi, dd);
                                      //System.Data.DataTable dt = SQLHelper.ExecuteRead(CommandType.Text, selSql, "aa");
                                      //if (dt.Rows.Count == 0)
                                      //{
                                      //    mut.WaitOne();
                                      //    string intsetsql = "Insert into HistoryGps([PDAID],[QQSJ],[YDSJ],[Lo],[La],[DevType]) values(" +
                                      //              " '" + issi + "', '" + dd.ToString() + "', '1900-01-01 00:00:00.000', '" + lat + "','" + lon + "','2')";

                                      //    SQLHelper.ExecuteNonQuery(intsetsql); 
                                      //    duiJiangJiRecive++;
                                      //    mut.ReleaseMutex();

                                      //}

                                      System.Threading.Thread.Sleep(60);
                                     

                }
                catch (Exception ex)
                {
                    WriteLog(ex.Message);
                }
            }
        }
        #endregion

        public void QueueInsert()
        {

            DataTable dt = GetTableSchema();
            bulkCopy.BatchSize = dt.Rows.Count;
            bulkCopy.ColumnMappings.Add(0, 3);
            bulkCopy.ColumnMappings.Add(1, 5);
            bulkCopy.ColumnMappings.Add(2, 6);
            bulkCopy.ColumnMappings.Add(3, 16);
            bulkCopy.ColumnMappings.Add(4, 17);

            bulkCopy.ColumnMappings.Add(5, 1);
            while (true)
            {
                try
                {
                    WriteLog("成功:" + System.DateTime.Now.ToString("F"));
                    for (int i = 0; i < gpsList.Count; i++)
                    {
                        GPSData data = gpsList[i];
                        DataRow dr = dt.NewRow();
                        //dr[0] = i;
                        dr[0] = data.PDAID;
                        dr[1] = data.QQSJ.ToString();
                        dr[2] = "1900-01-01 00:00:00.000";
                        //  dr[1] = DateTime.Parse(data.QQSJ.ToString());
                        //  dr[2] =  DateTime.Parse("1900-01-01 00:00:00.000");
                        dr[3] = data.LA;
                        dr[4] = data.LO;
                        dr[5] = 2;
                        dt.Rows.Add(dr);
                    }
                    duiJiangJiRecive += gpsList.Count;
                    gpsList.Clear();
                    if (dt != null && dt.Rows.Count != 0)
                    {
                        bulkCopy.WriteToServer(dt);
                        dt.Clear();
                    }

                    System.Threading.Thread.Sleep(1000 * 10);
                }
                catch (Exception ex)
                {
                    WriteLog("QueueInsert error:" + ex.Message);
                }
            }


            return;

            int xh = 1;
            while (xh == 1)
            {
                try
                {
                   
                    for (int i = 0; i < gpsList.Count; i++)
                    {
                        GPSData data = gpsList[i];
                        //mut.WaitOne();
                        string intsetsql = "Insert into HistoryGps_Temp([PDAID],[QQSJ],[YDSJ],[Lo],[La],[DevType]) values(" +
                                  " '" + data.PDAID + "', '" + data.QQSJ.ToString() + "', '1900-01-01 00:00:00.000', '" + data.LO + "','" + data.LA + "','2')";
                        //WriteLog(intsetsql);
                        ExecuteNonQuery(intsetsql);

                        //mut.ReleaseMutex();
                        duiJiangJiRecive++;
                        gpsList.Remove(data);
                        //WriteLog(gpsList.Count.ToString());
                        System.Threading.Thread.Sleep(50);
                    }
                }
                catch (Exception ex)
                {
                    WriteLog("QueueInsert error:"+ex.Message);
                }
            }
        }

        static DataTable GetTableSchema()
        {
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[] { 
      
      
        new DataColumn("PDAID",typeof(string)),  
        new DataColumn("QQSJ",typeof(DateTime)),  
        new DataColumn("YDSJ",typeof(DateTime)), 
        new DataColumn("Lo",typeof(decimal)), 
        new DataColumn("La",typeof(decimal)),
        new DataColumn("DevType",typeof(int))
         });
            return dt;
        }


        public Object ExecuteScalar(String sql)
        {
            using (SqlCommand cmd = new SqlCommand())
            {
                cmd.CommandText = sql;
                cmd.Connection = conn;
                if (cmd.Connection.State != ConnectionState.Open)
                    cmd.Connection.Open();
                return cmd.ExecuteScalar();
            }
        }

        public int ExecuteNonQuery(string cmdText)
        {
            using (SqlConnection conn = new SqlConnection(Program.sqladb))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = cmdText;
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = conn;
                    return cmd.ExecuteNonQuery();
                }
            }

        }
        public DataTable ExecuteRead(CommandType cmdtype, string cmdText, string tableName, params SqlParameter[] Parameters)
        {
            using (SqlConnection conn1 = new SqlConnection(Program.sqladb))
            {
                conn1.Open();
                SqlDataAdapter dr = new SqlDataAdapter(cmdText, conn1);
                DataSet ds = new DataSet();
                foreach (SqlParameter Parameter in Parameters)
                {
                    dr.SelectCommand.Parameters.Add(Parameter);
                }
                dr.Fill(ds, tableName);
                return ds.Tables[0];
            }

        }

        //public void DuiJiangJiDataHandler()
        //{
        //    UdpClient ReceiveClient;
        //    IPEndPoint remote = new IPEndPoint(IPAddress.Any, 0);
        //    ReceiveClient = new UdpClient(6800);
        //    int xh = 1;

        //    try
        //    {
        //        while (xh == 1)
        //        {
        //            //WriteLog("写入对讲机");
        //            //Console.WriteLine("写入对讲机.");


        //            byte[] bytes = null;
        //            // GPSData data;

        //            bytes = ReceiveClient.Receive(ref remote);

        //            //WriteLog("bytes vales:" + ByteToString(bytes));

        //            //string strBytes = "AA AA CC CC 03 00 35 00 00 00 35 37 36 31 31 32 30 31 00 00 00 00 00 00 00 00 00 00 00 00 18 44 A4 A6 5D 5C 5E 40 E1 25 38 F5 81 A8 3C 40 00 00 20 00 00 00 01 00 E1 07 08 04 0B 04 15 00 00";
        //            //strBytes = strBytes.Replace(" ", "");
        //            //bytes = DuiJiangJi.HexStrTobyte(strBytes);

        //            string issi = System.Text.Encoding.ASCII.GetString(bytes, 10, 20);
        //            issi = issi.Replace("\0", "");
        //            //经度
        //            double lon = 0.00000;
        //            Int32 size = Marshal.SizeOf(typeof(double));
        //            IntPtr buffer = Marshal.AllocHGlobal(size);
        //            try
        //            {
        //                Marshal.Copy(DuiJiangJi.getBbytes(bytes, 30, 8), 0, buffer, size);
        //                lon = (double)Marshal.PtrToStructure(buffer, typeof(double));
        //            }
        //            catch (Exception ex)
        //            {
        //                WriteLog(ex.Message);
        //            }
        //            finally
        //            {
        //                Marshal.FreeHGlobal(buffer);
        //                //GC.Collect();
        //            }

        //            //纬度
        //            double lat = 0.00000;
        //            size = Marshal.SizeOf(typeof(double));
        //            buffer = Marshal.AllocHGlobal(size);
        //            try
        //            {
        //                Marshal.Copy(DuiJiangJi.getBbytes(bytes, 38, 8), 0, buffer, size);
        //                lat = (double)Marshal.PtrToStructure(buffer, typeof(double));
        //            }
        //            catch (Exception ex)
        //            {
        //                WriteLog(ex.Message);
        //            }
        //            finally
        //            {
        //                Marshal.FreeHGlobal(buffer);
        //                //GC.Collect();
        //            }

        //            //年
        //            ushort year = 1900;
        //            size = Marshal.SizeOf(typeof(ushort));
        //            buffer = Marshal.AllocHGlobal(size);
        //            try
        //            {
        //                Marshal.Copy(DuiJiangJi.getBbytes(bytes, 54, 2), 0, buffer, size);
        //                year = (ushort)Marshal.PtrToStructure(buffer, typeof(ushort));
        //            }
        //            catch (Exception ex)
        //            {
        //                WriteLog(ex.Message);
        //            }
        //            finally
        //            {
        //                Marshal.FreeHGlobal(buffer);
        //                //GC.Collect();
        //            }

        //            string time = year + "-" + DuiJiangJi.getBbytes(bytes, 56, 1)[0].ToString() + "-" + DuiJiangJi.getBbytes(bytes, 57, 1)[0].ToString() +
        //        " " + DuiJiangJi.getBbytes(bytes, 58, 1)[0].ToString() + ":" + DuiJiangJi.getBbytes(bytes, 59, 1)[0].ToString() + ":" + DuiJiangJi.getBbytes(bytes, 60, 1)[0].ToString();
        //            string dd = DateTime.Parse(time).ToString("yyyy-MM-dd hh:mm:ss");
        //            duiJiangJiRecive1++;
        //            string selSql = string.Format("select * from [JingWuTong].[dbo].[HistoryGps] where PDAID='{0}' and QQSJ='{1}'", issi, System.DateTime.Parse(dd));
        //            System.Data.DataTable dt = SQLHelper.ExecuteRead(CommandType.Text, selSql, "aa");
        //            if (dt.Rows.Count == 0)
        //            {
        //                mut.WaitOne();
        //                string intsetsql = "Insert into HistoryGps([PDAID],[QQSJ],[YDSJ],[Lo],[La],[DevType]) values(" +
        //                          " '" + issi + "', '" + dd + "', '1900-01-01 00:00:00.000', '" + lat + "','" + lon + "','2')";

        //                SQLHelper.ExecuteNonQuery(intsetsql);
        //                mut.ReleaseMutex();
        //                duiJiangJiRecive++;
        //            }

        //            System.Threading.Thread.Sleep(50);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex.Message);
        //    }

        //}

        private void showTimer_Tick(object sender, EventArgs e)
        {
            this.Invoke((EventHandler)delegate
            {
                textBox1.Text = "";
                string strHistoryGps = "历史GPS信息：" + this.historyGpsRevice.ToString() + "条";
                //string strRealTimeGps = "\r\n实时信息：" + this.GpsRevice.ToString() + "条";
                string violation = "\r\n违法、妨碍：" + this.violationRevice.ToString() + "条";
                string force = "\r\n强制措施：" + this.forceRevice.ToString() + "条";
                string dev = "\r\n设备：" + this.devRevice.ToString() + "条";
                string duiJiangJi = "\r\n对讲机收到条数：" + duiJiangJiRecive1 + "条;入库条数：" + this.duiJiangJiRecive.ToString() + "条";
                textBox1.Text = strHistoryGps + violation + force + dev + duiJiangJi;

            });
        }

        private void ycsz_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RemoteConfig remoteForm = new RemoteConfig();
            remoteForm.Show();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //ReceiveThread.Abort();
            //threadStatus = ThreadState.Stopped;
            gpsList.Clear();
            if (ReceiveClient != null)
            {
                ReceiveClient.Close();
            }
            this.Dispose();
            System.Windows.Forms.Application.Exit();
            System.Environment.Exit(0);
        }

        private void WriteLog(string message)
        {
            string logFilename = "log" + DateTime.Now.ToString("yyMMdd") + ".txt";
            string logPath = Application.StartupPath + "\\" + logFilename;
            if (!File.Exists(logPath))
            {

                string str = "message:" + message + ",time:" + System.DateTime.Now.ToString() + "\r\n";
                BufferedStream bs = new BufferedStream(File.Create(logPath));
                byte[] info = new UTF8Encoding().GetBytes(str);
                bs.Write(info, 0, info.Length);
                bs.Close();

            }
            else
            {
                Encoding encoder = Encoding.UTF8;
                byte[] bytes = encoder.GetBytes("message:" + message + ",time:" + System.DateTime.Now.ToString() + "\r\n");
                FileStream fs = File.OpenWrite(logPath);
                //设定书写的开始位置为文件的末尾
                fs.Position = fs.Length;
                //将待写入内容追加到文件末尾
                fs.Write(bytes, 0, bytes.Length);
                fs.Close();


            }


        }


    }
}
