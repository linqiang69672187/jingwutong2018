using DbComponent;
using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Web;

namespace JingWuTong.Handle
{
    /// <summary>
    /// exportAll_datamanagement 的摘要说明
    /// </summary>
    public class exportAll_datamanagement : IHttpHandler
    {

        DataTable allEntitys = null;  //递归单位信息表
        List<dataStruct> tmpList = new List<dataStruct>();
        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            string type = context.Request.Form["type"];
            string begintime = context.Request.Form["begintime"];
            string endtime = context.Request.Form["endtime"];
            string hbbegintime = context.Request.Form["hbbegintime"];
            string hbendtime = context.Request.Form["hbendtime"];
            string ssdd = context.Request.Form["ssdd"];
            string sszd = context.Request.Form["sszd"];
            string requesttype = context.Request.Form["requesttype"];
            string search = context.Request.Form["search"];
            string sreachcondi = "";

            int onlinevalue = int.Parse(context.Request.Form["onlinevalue"]) * 60;
            int usedvalue = int.Parse(context.Request.Form["usedvalue"]) * 60;


            if (search != "")
            {
               // sreachcondi = " (de.[DevId] like '%" + search + "%' or us.[XM] like '%" + search + "%' or us.[JYBH] like '%" + search + "%' ) and ";
            }

            string tmpDevid = "";
            int tmpRows = 0;
            DataTable dtEntity = null;  //单位信息表

            DataTable Alarm_EveryDayInfo = null; //每日告警
            DataTable dUser = null;
     



            StringBuilder sqltext = new StringBuilder();

            DataTable dtreturns = new DataTable(); //返回数据表
            dtreturns.Columns.Add("cloum1");
            dtreturns.Columns.Add("cloum2");
            dtreturns.Columns.Add("cloum3");
            dtreturns.Columns.Add("cloum4");
            dtreturns.Columns.Add("cloum5");
            dtreturns.Columns.Add("cloum6", typeof(double));
            dtreturns.Columns.Add("cloum7", typeof(double));
            dtreturns.Columns.Add("cloum8");
            dtreturns.Columns.Add("cloum9");
            dtreturns.Columns.Add("cloum10");
            dtreturns.Columns.Add("cloum11");
            dtreturns.Columns.Add("cloum12");
            dtreturns.Columns.Add("cloum13", typeof(int));
            dtreturns.Columns.Add("cloum14");



            int days = Convert.ToInt16(context.Request.Form["dates"]);
            int statusvalue = 10;  //正常参考值
            int zxstatusvalue = 30;//在线参考值
            int devicescount = 0;  //汇总设备总数
            double zxsc = 0.0;  //汇总在线时长
            double spdx = 0.0;  //汇总视频大小
            Int64 cxl = 0;  //汇总查询量
            Int64 jwtzxsc = 0;  //警务通在线时长
            int hzusecount = 0;
            int wcxl = 0;   //无查询量设备数量
            int wcfl = 0;   //无处罚量设备数量
            int wsysb = 0;  //无使用设备数量
            int zxsb = 0; //在线设备
            string pxstring = "";
            string bmdm = ""; //汇总的部门代码
            int allstatu_device = 0;  //汇总使用率不为空数量
            string ddtitle;//大队标题


            statusvalue = days * usedvalue;//超过10分钟算使用
            zxstatusvalue = days * onlinevalue;//在线参考值

            allEntitys = SQLHelper.ExecuteRead(CommandType.Text, "SELECT BMDM,SJBM,BMMC,Sort from [Entity] ", "11");
            DataTable devtypes = SQLHelper.ExecuteRead(CommandType.Text, "SELECT TypeName,ID FROM [dbo].[DeviceType] where ID<7  ORDER by Sort ", "11");
            dUser = SQLHelper.ExecuteRead(CommandType.Text, "SELECT en.SJBM,us.BMDM FROM [dbo].[ACL_USER] us left join Entity en on us.BMDM = en.BMDM", "user");
            DataTable zfData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT  sum(CONVERT(bigint,[VideLength])) as 视频长度, sum(CONVERT(bigint,[FileSize])) as 文件大小,sum([UploadCnt]) as 上传量,sum([GFUploadCnt]) as 规范上传量,de.BMDM,de.DevId FROM [EveryDayInfo_ZFJLY] al left join Device de on de.DevId = al.DevId where  [Time] >='" + begintime + "' and [Time] <='" + endtime + "'  group by de.DevId,de.BMDM", "Alarm_EveryDayInfo");
            DataTable zxscData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT de.BMDM,SUM(value) as value ,de.DevId  FROM [Alarm_EveryDayInfo] al left join Device de on de.DevId = al.DevId where al.AlarmType=1 group by de.DevId,de.BMDM ", "Alarm_EveryDayInfo");
            DataTable cllData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT de.BMDM,SUM(value) as  value ,de.DevId  FROM [Alarm_EveryDayInfo] al left join Device de on de.DevId = al.DevId where al.AlarmType=2 group by de.DevId,de.BMDM ", "Alarm_EveryDayInfo");
            DataTable cxlData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT de.BMDM,SUM(value) as  value ,de.DevId  FROM [Alarm_EveryDayInfo] al left join Device de on de.DevId = al.DevId where al.AlarmType=5 group by de.DevId,de.BMDM ", "Alarm_EveryDayInfo");


            //所有大队

            for (int h = 0; h < devtypes.Rows.Count; h++)
            {
                var rows = from p in allEntitys.AsEnumerable()
                           where (p.Field<string>("BMDM")== "331000000000")
                           orderby p.Field<double>("Sort") descending
                           select p;
                    foreach (var entityitem in rows)
                    {
                        if (devtypes.Rows[h]["TypeName"].ToString() != "执法记录仪") continue;




                    }





                }

               


        

            //string reTitle = ExportExcel(dtreturns, type, begintime, endtime, ssdd, sszd);
          
        }

        public IEnumerable<entityStruct> GetSonID(string p_id)
        {
            try
            {
                var query = (from p in allEntitys.AsEnumerable()
                             where (p.Field<string>("SJBM") == p_id)
                             select new entityStruct
                             {
                                 BMDM = p.Field<string>("BMDM"),
                                 SJBM = p.Field<string>("SJBM")
                             }).ToList<entityStruct>();
                return query.ToList().Concat(query.ToList().SelectMany(t => GetSonID(t.BMDM)));
            }
            catch (Exception e)
            {
                return null;
            }

        }

        public List<dataStruct> findallchildren(string parentid, DataTable dt)
        {
            var list = (from p in dt.AsEnumerable()
                        where p.Field<string>("ParentID") == parentid
                        select new dataStruct
                        {
                            BMDM = p.Field<string>("BMDM"),
                            ParentID = p.Field<string>("ParentID"),
                            在线时长 = p.Field<int>("在线时长"),
                            文件大小 = p.Field<int>("文件大小"),
                            AlarmType = p.Field<int>("AlarmType"),
                            DevId = p.Field<string>("DevId")
                        }).ToList<dataStruct>();
            if (list.Count != 0)
            {
                tmpList.AddRange(list);
            }
            foreach (dataStruct single in list)
            {
                List<dataStruct> tmpChildren = findallchildren(single.BMDM, dt);

            }
            return tmpList;
        }

        public class entityStruct
        {
            public string BMDM;
            public string SJBM;
        }

        public class dataStruct
        {
            public string BMDM = "BMDM";
            public string ParentID = "ParentID";
            public int 在线时长 = 0;
            public int 文件大小 = 0;
            public int AlarmType = 0;
            public string DevId = "DevId";
        }


        public string testExportExcel(IEnumerable<dataStruct> rows)
        {
            ExcelFile excelFile = new ExcelFile();
            var tmpath = "";
            tmpath = HttpContext.Current.Server.MapPath("templet\\xx.xls");


            excelFile.LoadXls(tmpath);
            ExcelWorksheet sheet = excelFile.Worksheets[0];

            int i = 0;


            foreach (dataStruct item in rows)
            {

                sheet.Rows[i + 2].Cells["A"].Value = item.BMDM;
                sheet.Rows[i + 2].Cells["B"].Value = item.DevId;
                sheet.Rows[i + 2].Cells["C"].Value = item.ParentID;
                sheet.Rows[i + 2].Cells["D"].Value = item.在线时长;
                sheet.Rows[i + 2].Cells["E"].Value = item.文件大小;
                sheet.Rows[i + 2].Cells["F"].Value = item.AlarmType;

                i += 1;

            }


            tmpath = HttpContext.Current.Server.MapPath("upload\\xx.xls");

            excelFile.SaveXls(tmpath);
            return sheet.Rows[0].Cells[0].Value + ".xls";
        }


        public string ExportExcel(DataTable dt, string type, string begintime, string endtime, string entityTitle, string ssdd, string sszd, string ssddtext, string sszdtext)
        {
            ExcelFile excelFile = new ExcelFile();
            var tmpath = "";
            string Entityname = "";
            Entityname += (ssddtext == "全部") ? "台州交警局" : ssddtext;
            Entityname += (sszdtext == "全部") ? "" : sszdtext;
            switch (type)
            {
                case "1":
                case "2":
                case "3":
                    tmpath = HttpContext.Current.Server.MapPath("templet\\1.xls");
                    break;
                case "5":
                    tmpath = HttpContext.Current.Server.MapPath("templet\\5.xls");
                    break;
                case "4":
                    tmpath = HttpContext.Current.Server.MapPath("templet\\4.xls");
                    break;
                case "6":
                    tmpath = HttpContext.Current.Server.MapPath("templet\\6.xls");
                    break;

            }

            excelFile.LoadXls(tmpath);
            ExcelWorksheet sheet = excelFile.Worksheets[0];


            DateTime bg = Convert.ToDateTime(begintime);
            DateTime ed = Convert.ToDateTime(endtime);
            string typename = "";
            switch (type)
            {
                case "1":
                    typename = "车载视频";
                    break;
                case "2":
                    typename = "对讲机";
                    break;
                case "3":
                    typename = "拦截仪";
                    break;
                case "5":
                    typename = "执法记录仪";
                    break;
                case "4":
                    typename = "警务通";
                    break;
                case "6":
                    typename = "辅警通";
                    break;
            }


            sheet.Rows[0].Cells["A"].Value = begintime.Replace("/", "-") + "_" + endtime.Replace("/", "-") + Entityname + typename + "报表";
            switch (type)
            {
                case "1":
                case "2":
                case "3":
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sheet.Rows[i + 2].Cells["A"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["A"].Value = dt.Rows[i][0].ToString();
                        sheet.Rows[i + 2].Cells["B"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["B"].Value = dt.Rows[i][1].ToString();
                        sheet.Rows[i + 2].Cells["C"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["C"].Value = dt.Rows[i][2].ToString();
                        sheet.Rows[i + 2].Cells["D"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["D"].Value = dt.Rows[i][4].ToString();
                        sheet.Rows[i + 2].Cells["E"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["E"].Value = dt.Rows[i][3].ToString();
                        sheet.Rows[i + 2].Cells["F"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["F"].Value = dt.Rows[i][6].ToString();
                        sheet.Rows[i + 2].Cells["G"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["G"].Value = dt.Rows[i][7].ToString();

                    }
                    break;

                case "5":
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sheet.Rows[i + 2].Cells["A"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["A"].Value = dt.Rows[i][0].ToString();
                        sheet.Rows[i + 2].Cells["B"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["B"].Value = dt.Rows[i][1].ToString();
                        sheet.Rows[i + 2].Cells["C"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["C"].Value = dt.Rows[i][2].ToString();
                        sheet.Rows[i + 2].Cells["D"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["D"].Value = dt.Rows[i][3].ToString();
                        sheet.Rows[i + 2].Cells["E"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["E"].Value = dt.Rows[i][8].ToString();
                        sheet.Rows[i + 2].Cells["F"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["F"].Value = dt.Rows[i][4].ToString();
                        sheet.Rows[i + 2].Cells["G"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["G"].Value = dt.Rows[i][6].ToString();
                        sheet.Rows[i + 2].Cells["H"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["H"].Value = dt.Rows[i][5].ToString();
                        sheet.Rows[i + 2].Cells["I"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["I"].Value = dt.Rows[i][7].ToString();
                    }

                    break;
                case "4":
                case "6":
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sheet.Rows[i + 2].Cells["A"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["A"].Value = dt.Rows[i][0].ToString();
                        sheet.Rows[i + 2].Cells["B"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["B"].Value = dt.Rows[i][1].ToString();
                        sheet.Rows[i + 2].Cells["C"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["C"].Value = dt.Rows[i][2].ToString();
                        sheet.Rows[i + 2].Cells["D"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["D"].Value = dt.Rows[i][3].ToString();
                        sheet.Rows[i + 2].Cells["E"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["E"].Value = dt.Rows[i][8].ToString();
                        sheet.Rows[i + 2].Cells["F"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["F"].Value = dt.Rows[i][4].ToString();
                        sheet.Rows[i + 2].Cells["G"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["G"].Value = dt.Rows[i][6].ToString();
                        sheet.Rows[i + 2].Cells["H"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["H"].Value = dt.Rows[i][5].ToString();
                        sheet.Rows[i + 2].Cells["I"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["I"].Value = dt.Rows[i][7].ToString();
                        sheet.Rows[i + 2].Cells["J"].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                        sheet.Rows[i + 2].Cells["J"].Value = dt.Rows[i][10].ToString();
                    }

                    break;

            }


            tmpath = HttpContext.Current.Server.MapPath("upload\\" + sheet.Rows[0].Cells[0].Value + ".xls");

            excelFile.SaveXls(tmpath);
            return sheet.Rows[0].Cells[0].Value + ".xls";
        }


        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}