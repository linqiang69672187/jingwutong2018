using DbComponent;
using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web;


namespace JingWuTong.Handle
{
    /// <summary>
    /// exportAll_Timesharing_Reports 的摘要说明
    /// </summary>
    public class exportAll_Timesharing_Reports : IHttpHandler
    {
        List<dataStruct> tmpList = new List<dataStruct>();

        DataTable allEntitys = null;
        DataTable devtypes = null;
        DataTable dUser = null;
        DataTable zfData = null;
        DataTable Data = null;



        int statusvalue = 0;  //正常参考值
        int zxstatusvalue = 0;//在线参考值

        int sheetrows = 0;
        int dataindex = 0;
        string begintime = "";
        string endtime = "";
        int countTime;
        int currentTime = 0;
        ExcelFile excelFile = null;
        private log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            string type = context.Request.Form["type"];
            begintime = context.Request.Form["begintime"];
            endtime = context.Request.Form["endtime"];
            string ssdd = context.Request.Form["ssdd"];
            string sszd = context.Request.Form["sszd"];
            string requesttype = context.Request.Form["requesttype"];
            string search = context.Request.Form["search"];
            string sreachcondi = "";

            int onlinevalue = int.Parse(context.Request.Form["onlinevalue"]) * 60;
            int usedvalue = int.Parse(context.Request.Form["usedvalue"]) * 60;


            if (search != "")
            {
                sreachcondi = " (de.[IMEI] like '%" + search + "%' or de.[DevId] like '%" + search + "%' or us.[XM] like '%" + search + "%' or us.[JYBH] like '%" + search + "%' ) and ";
            }

            foreach (var key in ConfigurationManager.AppSettings.AllKeys)
            {
                countTime += (key.Contains("Time")) ? 1 : 0;
            }



            allEntitys = SQLHelper.ExecuteRead(CommandType.Text, "SELECT BMDM,SJBM,BMQC as BMMC,isnull(Sort,0) as Sort,id from [Entity] ", "11");
            devtypes = SQLHelper.ExecuteRead(CommandType.Text, "SELECT TypeName,ID FROM [dbo].[DeviceType] where ID<7  ORDER by Sort ", "11");
            log.Info("读取zfData表");

            zfData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT VideLength, [FileSize] ,[UploadCnt],[GFUploadCnt],de.BMDM,de.DevId,substring(convert(varchar,[Time],120),12,5) Time FROM [EveryDayInfo_ZFJLY_Hour] al left join Device de on de.DevId = al.DevId  left join ACL_USER as us on de.JYBH = us.JYBH     where " + sreachcondi + "   [Time] >='" + begintime + "' and [Time] <='" + endtime + " 23:59' and de.devType='5' ", "Alarm_EveryDayInfo");
            log.Info("表zfData完成");

            dUser = SQLHelper.ExecuteRead(CommandType.Text, "SELECT en.SJBM,us.BMDM,us.XM FROM [dbo].[ACL_USER] us  left join  Entity en  on us.BMDM = en.BMDM ", "user");
            log.Info("读取Data表");
            Data = SQLHelper.ExecuteRead(CommandType.Text, "SELECT OnlineTime, [HandleCnt] ,[CXCnt],de.BMDM,de.DevId,substring(convert(varchar,[Time],120),12,5) Time,de.devtype FROM [EverydayInfo_Hour] al left join Device de on de.DevId = al.DevId  left join ACL_USER as us on de.JYBH = us.JYBH     where " + sreachcondi + "   [Time] >='" + begintime + "' and [Time] <='" + endtime + " 23:59' and de.devType in (1,2,3,4,6)", "Alarm_EveryDayInfo");
            log.Info("表Data完成");







            int days = Convert.ToInt16(context.Request.Form["dates"]);



            statusvalue = days * usedvalue;//超过10分钟算使用
            zxstatusvalue = days * onlinevalue;//在线参考值


             excelFile = new ExcelFile();
            var tmpath = "";
            tmpath = HttpContext.Current.Server.MapPath("templet\\0.xls");
            excelFile.LoadXls(tmpath);
            //所有大队

            for (int h = 0; h < devtypes.Rows.Count; h++)
            {

                string typename = devtypes.Rows[h]["TypeName"].ToString();
                Thread thread = new Thread(new ParameterizedThreadStart(ThreadInsertSheet));
                thread.Start(typename);
                log.Info(typename + "_线程开始");



            }

            while (true)
            {
                Thread.Sleep(1000);
                if (currentTime == devtypes.Rows.Count)
                {
                    tmpath = HttpContext.Current.Server.MapPath("upload\\" + begintime.Replace("/", "-") + "_" + endtime.Replace("/", "-") + "分时段报表.xls");
                    excelFile.SaveXls(tmpath);
                    StringBuilder retJson = new StringBuilder();


                    retJson.Append("{\"");
                    retJson.Append("data");
                    retJson.Append('"');
                    retJson.Append(":");
                    retJson.Append('"');
                    retJson.Append(begintime.Replace("/", "-") + "_" + endtime.Replace("/", "-") + "分时段报表.xls");
                    retJson.Append('"');
                    retJson.Append("}");
                    context.Response.Write(retJson);
                    return;
                    //string reTitle = ExportExcel(dtreturns, type, begintime, endtime, ssdd, sszd);
                }
            }

        }


        public void ThreadInsertSheet(object typename)
        {
            ExcelWorksheet sheet = excelFile.Worksheets[typename.ToString()];

            string typeid = "0";
            switch (typename.ToString())
            {
                case "车载视频":
                    typeid = "1";
                    break;
                case "对讲机":
                    typeid = "2";
                    break;
                case "拦截仪":
                    typeid = "3";
                    break;
                case "警务通":
                    typeid = "4";
                    break;
                case "执法记录仪":
                    typeid = "5";
                    break;
                case "辅警通":
                    typeid = "6";
                    break;
                case "测速仪":
                    typeid = "7";
                    break;
                case "酒精测试仪":
                    typeid = "8";
                    break;

            }
            try
            {
                InsertRowdata(sheet, typeid, typename.ToString(), "331000000000", "支队", "台州市交通警察局");
            }
            catch (Exception e)
            {
                sheet.Rows[0].Cells["A"].Value = e.ToString();

            }
            currentTime += 1;
            log.Info(typename + "_线程结束");

        }


        public CellStyle Titlestyle()
        {
            CellStyle style = new CellStyle();
            //设置水平对齐模式
            style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            //设置垂直对齐模式
            style.VerticalAlignment = VerticalAlignmentStyle.Center;
            //设置字体
            style.Font.Size = 12 * 20; //PT=20
            style.Font.Weight = ExcelFont.BoldWeight;
            style.FillPattern.SetPattern(FillPatternStyle.Solid,ColorTranslator.FromHtml("#ccffcc"), Color.Empty);
            style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
            //  style.Font.Color = Color.Blue;
            return style;
        }

        public void InsertTitle(ExcelWorksheet sheet, string type)
        {
            CellRange range;
            CellRange rangebm;
            CellRange rangepf;
            CellRange rangejy;
            CellStyle style;
            int mergedint = 0;
            int h = 0;
            switch (type)
            {
                case "1":
                case "2":
                case "3":
                    mergedint = 1 + countTime * 3;
                    range = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows+1, mergedint);
                    style = new CellStyle();
                    style.FillPattern.SetPattern(FillPatternStyle.Solid, ColorTranslator.FromHtml("#ccffcc"), Color.Empty);
                    style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                    style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                    //设置垂直对齐模式
                    style.Font.Size = 12 * 20; //PT=20
                    style.Font.Weight = ExcelFont.BoldWeight;
                    style.VerticalAlignment = VerticalAlignmentStyle.Center;
                    range.Style = style;

                    rangebm = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows+1, 0);//GetSubrange("A1", "G1");
                    rangepf = sheet.Cells.GetSubrangeAbsolute(sheetrows, 1, sheetrows + 1, 1);//GetSubrange("A1", "G1");
                    rangebm.Value = "部门";
                    rangepf.Value = "设备配发数（台）";
                    rangebm.Merged = true;
                    rangepf.Merged = true;
                    
                    foreach (var key in ConfigurationManager.AppSettings.AllKeys)
                    {
                        if (!key.Contains("Time")) continue;
                        CellRange timerange = sheet.Cells.GetSubrangeAbsolute(sheetrows, 2+h, sheetrows , 4 + h );
                        timerange.Merged = true;
                        timerange.Value = ConfigurationManager.AppSettings[key];
                        sheet.Rows[sheetrows+1].Cells[2 + h].Value = "设备使用数量（台）";
                        sheet.Rows[sheetrows+1].Cells[3 + h].Value = "在线时长总和(小时)";
                        sheet.Rows[sheetrows+1].Cells[4 + h].Value = "设备使用率";
                        h += 3;
                    }

                    sheetrows += 1;
                    //      range.Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                    break;
                case "4":
                    mergedint = 2 + countTime * 5;
                    range = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows + 1, mergedint);
                    style = new CellStyle();
                    style.FillPattern.SetPattern(FillPatternStyle.Solid, ColorTranslator.FromHtml("#ccffcc"), Color.Empty);
                    style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                    style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                    //设置垂直对齐模式
                    style.Font.Size = 12 * 20; //PT=20
                    style.Font.Weight = ExcelFont.BoldWeight;
                    style.VerticalAlignment = VerticalAlignmentStyle.Center;
                    range.Style = style;

                    rangebm = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows + 1, 0);//GetSubrange("A1", "G1");
                    rangepf = sheet.Cells.GetSubrangeAbsolute(sheetrows, 1, sheetrows + 1, 1);//GetSubrange("A1", "G1");
                    rangejy = sheet.Cells.GetSubrangeAbsolute(sheetrows, 2, sheetrows + 1, 2);//GetSubrange("A1", "G1");
                    rangebm.Value = "部门";
                    rangepf.Value = "配发数量（台）";
                    rangejy.Value = "警员数";
                    rangebm.Merged = true;
                    rangepf.Merged = true;
                    rangejy.Merged = true;
                    foreach (var key in ConfigurationManager.AppSettings.AllKeys)
                    {
                        if (!key.Contains("Time")) continue;
                        CellRange timerange = sheet.Cells.GetSubrangeAbsolute(sheetrows, 3 + h, sheetrows, 7 + h);
                        timerange.Merged = true;
                        timerange.Value = ConfigurationManager.AppSettings[key];
                        sheet.Rows[sheetrows + 1].Cells[3 + h].Value = "警务通处罚数（例）";
                        sheet.Rows[sheetrows + 1].Cells[4 + h].Value = "人均处罚量";
                        sheet.Rows[sheetrows + 1].Cells[5 + h].Value = "查询量（次）";
                        sheet.Rows[sheetrows + 1].Cells[6 + h].Value = "设备平均处罚量";
                        sheet.Rows[sheetrows + 1].Cells[7 + h].Value = "无处罚数的警务通（台）";
                        h += 5;
                    }

                    sheetrows += 1;
                    break;
                case "6":
                    mergedint = 2 + countTime * 5;
                    range = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows + 1, mergedint);
                    style = new CellStyle();
                    style.Font.Size = 12 * 20; //PT=20
                    style.Font.Weight = ExcelFont.BoldWeight;
                    style.FillPattern.SetPattern(FillPatternStyle.Solid, ColorTranslator.FromHtml("#ccffcc"), Color.Empty);
                    style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                    style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                    //设置垂直对齐模式
                    style.VerticalAlignment = VerticalAlignmentStyle.Center;
                    range.Style = style;

                    rangebm = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows + 1, 0);//GetSubrange("A1", "G1");
                    rangepf = sheet.Cells.GetSubrangeAbsolute(sheetrows, 1, sheetrows + 1, 1);//GetSubrange("A1", "G1");
                    rangejy = sheet.Cells.GetSubrangeAbsolute(sheetrows, 2, sheetrows + 1, 2);//GetSubrange("A1", "G1");
                    rangebm.Value = "部门";
                    rangepf.Value = "配发数量（台）";
                    rangejy.Value = "辅警数";
                    rangebm.Merged = true;
                    rangepf.Merged = true;
                    rangejy.Merged = true;
                    foreach (var key in ConfigurationManager.AppSettings.AllKeys)
                    {
                        if (!key.Contains("Time")) continue;
                        CellRange timerange = sheet.Cells.GetSubrangeAbsolute(sheetrows, 3 + h, sheetrows, 7 + h);
                        timerange.Merged = true;
                        timerange.Value = ConfigurationManager.AppSettings[key];
                        sheet.Rows[sheetrows + 1].Cells[3 + h].Value = "违停采集（例）";
                        sheet.Rows[sheetrows + 1].Cells[4 + h].Value = "人均处罚量";
                        sheet.Rows[sheetrows + 1].Cells[5 + h].Value = "查询量（次）";
                        sheet.Rows[sheetrows + 1].Cells[6 + h].Value = "设备平均处罚量";
                        sheet.Rows[sheetrows + 1].Cells[7 + h].Value = "无违停设备（台）";
                        h += 5;
                    }

                    sheetrows += 1;
                    break;
                case "5":
                    mergedint =1 + countTime * 6;
                    range = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows + 1, mergedint);
                    style = new CellStyle();
                    style.Font.Size = 12 * 20; //PT=20
                    style.Font.Weight = ExcelFont.BoldWeight;
                    style.FillPattern.SetPattern(FillPatternStyle.Solid, ColorTranslator.FromHtml("#ccffcc"), Color.Empty);
                    style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                    style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                    //设置垂直对齐模式
                    style.VerticalAlignment = VerticalAlignmentStyle.Center;
                    range.Style = style;

                    rangebm = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows + 1, 0);//GetSubrange("A1", "G1");
                    rangepf = sheet.Cells.GetSubrangeAbsolute(sheetrows, 1, sheetrows + 1, 1);//GetSubrange("A1", "G1");
                    rangebm.Value = "部门";
                    rangepf.Value = "设备配发数(台)";
                    rangebm.Merged = true;
                    rangepf.Merged = true;
                    foreach (var key in ConfigurationManager.AppSettings.AllKeys)
                    {
                        if (!key.Contains("Time")) continue;
                        CellRange timerange = sheet.Cells.GetSubrangeAbsolute(sheetrows, 2 + h, sheetrows, 7 + h);
                        timerange.Merged = true;
                        timerange.Value = ConfigurationManager.AppSettings[key];
                        sheet.Rows[sheetrows + 1].Cells[2 + h].Value = "设备使用数量（台）";
                        sheet.Rows[sheetrows + 1].Cells[3 + h].Value = "设备未使用数量（台";
                        sheet.Rows[sheetrows + 1].Cells[4 + h].Value = "视频时长总和（小时）";
                        sheet.Rows[sheetrows + 1].Cells[5 + h].Value = "视频大小(GB)";
                        sheet.Rows[sheetrows + 1].Cells[6 + h].Value = "48小时规范上传率";
                        sheet.Rows[sheetrows + 1].Cells[7 + h].Value = "设备使用率";
                        h += 6;
                    }

                    sheetrows += 1;
                    break;

            }
            sheetrows += 1;
        }

        public void InsertRowdata(ExcelWorksheet sheet, string type, string typename, string sjbm, string reporttype, string title)
        {
            DataTable dtreturns = new DataTable(); //返回数据表
            DataRow drtz = dtreturns.NewRow();

            int mergedint = 0;
            switch (type)
            {
                case "1":
                case "2":
                case "3":
                    mergedint = 1 + countTime * 3;
                    break;
                case "4":
                case "6":
                    mergedint = 2 + countTime * 5;
                    break;
                case "5":
                    mergedint = 1 + countTime * 6;
                    break;
            }

            for (int h = 0; h <= mergedint; h++)
            {
                dtreturns.Columns.Add(h.ToString());
            }

            dataindex = 0;
            string pxstring = "";
            OrderedEnumerableRowCollection<DataRow> rows;
            if (reporttype == "支队")
            {
                rows = from p in allEntitys.AsEnumerable()
                       where (p.Field<string>("SJBM") == sjbm)
                       orderby p.Field<int>("Sort") descending
                       select p;
            }
            else
            {
                rows = from p in allEntitys.AsEnumerable()
                       where (p.Field<string>("SJBM") == sjbm || p.Field<string>("BMDM") == sjbm)
                       orderby p.Field<int>("Sort") descending
                       select p;
            }

            foreach (var entityitem in rows)
            {
                if (type != "5" && entityitem["BMDM"].ToString() == "33100000000x") continue;//如果不是执法记录仪，跳出“局机关”单位
                DataRow dr = dtreturns.NewRow();
                dataindex += 1;
                dr["0"] = entityitem["BMMC"].ToString();  //部门名称

                var entityids = GetSonID(entityitem["BMDM"].ToString());
                List<string> strList = new List<string>();
                strList.Add(entityitem["BMDM"].ToString());
                if (!(reporttype != "支队" && entityitem["SJBM"].ToString() == "331000000000"))  //非支队报表下的大队单位，只显示本级
                {
                    foreach (entityStruct item in entityids)
                    {
                        strList.Add(item.BMDM);
                    }
                }
                List<dataStruct> queryrows;
                int h = 0;
                switch (type)
                {
                    case "1":
                    case "2":
                    case "3":
                        try {
                        foreach (var key in ConfigurationManager.AppSettings.AllKeys)
                        {
                            if (!key.Contains("Time")) continue;

                            int Ftime = int.Parse(ConfigurationManager.AppSettings[key].Split('-')[0].Replace(":",""));
                            int Stime = int.Parse(ConfigurationManager.AppSettings[key].Split('-')[1].Replace(":", ""));
                            int usedevices = 0;
                            Int64 onlinetime = 0;
                            queryrows = (from p in Data.AsEnumerable()
                                         where strList.ToArray().Contains(p.Field<string>("BMDM")) && p.Field<int>("devtype") == int.Parse(type) && int.Parse(p.Field<string>("Time").Replace(":",""))>= Ftime && int.Parse(p.Field<string>("Time").Replace(":", "")) < Stime
                                         group p by new
                                         {
                                             t1 = p.Field<string>("devid")
                                        

                                         } into g
                                         select new dataStruct
                                         {
                                             在线时长 = g.Sum(p => p.Field<int>("OnlineTime"))
                                         }).ToList<dataStruct>();
                                foreach (dataStruct item in queryrows)
                                {
                                    onlinetime += item.在线时长;
                                    usedevices += ((item.在线时长) - statusvalue > 0) ? 1 : 0;
                                }
                                dr[2 + h] = usedevices;
                                drtz[2 + h] = (drtz[2 + h].ToString() == "") ? usedevices : int.Parse(drtz[2 + h].ToString()) + usedevices; 
                                dr[3 + h] = Math.Round((double)onlinetime / 3600, 2);
                                drtz[3 + h] = (drtz[3 + h].ToString() == "") ? Math.Round((double)onlinetime / 3600, 2) : double.Parse(drtz[3 + h].ToString()) + Math.Round((double)onlinetime / 3600, 2);
                                dr[4 + h] =(queryrows.Count==0)?0:Math.Round((double)usedevices *100/ queryrows.Count, 2); ;
                                if (h == 0) {
                                    dr["1"] = queryrows.Count;
                                    drtz["1"] =(drtz["1"].ToString()=="")? queryrows.Count: int.Parse(drtz["1"].ToString())+queryrows.Count;
                                }

                            h += 3;

                        }
                        }
                        catch (Exception e) {

                        }
                        break;
                    case "4":
                    case "6":
                        try
                        {
                            var userrow = (from p in dUser.AsEnumerable()
                                           where strList.ToArray().Contains(p.Field<string>("BMDM"))
                                           select p);
                            foreach (var key in ConfigurationManager.AppSettings.AllKeys)
                            {
                                if (!key.Contains("Time")) continue;

                                int Ftime = int.Parse(ConfigurationManager.AppSettings[key].Split('-')[0].Replace(":", ""));
                                int Stime = int.Parse(ConfigurationManager.AppSettings[key].Split('-')[1].Replace(":", ""));
                                queryrows = (from p in Data.AsEnumerable()
                                             where strList.ToArray().Contains(p.Field<string>("BMDM")) && p.Field<int>("devtype") == int.Parse(type) && int.Parse(p.Field<string>("Time").Replace(":", "")) >= Ftime && int.Parse(p.Field<string>("Time").Replace(":", "")) < Stime
                                             group p by new
                                             {
                                                 t1 = p.Field<string>("devid")


                                             } into g
                                             select new dataStruct
                                             {
                                                 HandleCnt = g.Sum(p => p.Field<int>("HandleCnt")),
                                                 CXCnt = g.Sum(p => p.Field<int>("CXCnt")),
                                             }).ToList<dataStruct>();
                                int HandleCnt = 0;
                                int CXCnt = 0;
                                int NoneHandleCnt = 0;
                                foreach (dataStruct item in queryrows)
                                {
                                    HandleCnt += item.HandleCnt;
                                    CXCnt += item.CXCnt;
                                    NoneHandleCnt += (item.HandleCnt == 0) ? 1 : 0;
                                }
                                dr[3 + h] = HandleCnt;
                                dr[4 + h] = (userrow.Count()==0)?0:Math.Round((double)HandleCnt / userrow.Count(), 2); 
                                dr[5 + h] = CXCnt;
                                dr[6 + h] = (queryrows.Count == 0) ? 0 : Math.Round((double)HandleCnt / queryrows.Count, 2); 
                                dr[7 + h] = NoneHandleCnt;
                                drtz[3 + h] = (drtz[3 + h].ToString() == "") ? HandleCnt : int.Parse(drtz[3 + h].ToString()) + HandleCnt;
                                drtz[5 + h] = (drtz[5 + h].ToString() == "") ? CXCnt : int.Parse(drtz[5 + h].ToString()) + CXCnt;
                                drtz[7 + h] = (drtz[7 + h].ToString() == "") ? NoneHandleCnt : int.Parse(drtz[7 + h].ToString()) + NoneHandleCnt;

                                if (h == 0)
                                {
                                    dr["1"] = queryrows.Count;
                                    dr["2"] = userrow.Count().ToString();
                                    drtz["1"] = (drtz["1"].ToString() == "") ? queryrows.Count : int.Parse(drtz["1"].ToString()) + queryrows.Count;
                                    drtz["2"] = (drtz["2"].ToString() == "") ? userrow.Count() : int.Parse(drtz["2"].ToString()) + userrow.Count();

                                }

                                h += 5;

                            }
                        }
                        catch (Exception e)
                        {

                        }
                        break;
                    case "5":
                        try
                        {
                            foreach (var key in ConfigurationManager.AppSettings.AllKeys)
                            {
                                if (!key.Contains("Time")) continue;

                                int Ftime = int.Parse(ConfigurationManager.AppSettings[key].Split('-')[0].Replace(":", ""));
                                int Stime = int.Parse(ConfigurationManager.AppSettings[key].Split('-')[1].Replace(":", ""));

                                queryrows = (from p in zfData.AsEnumerable()
                                             where strList.ToArray().Contains(p.Field<string>("BMDM")) && int.Parse(p.Field<string>("Time").Replace(":", "")) >= Ftime && int.Parse(p.Field<string>("Time").Replace(":", "")) < Stime
                                             group p by new
                                             {
                                                 t1 = p.Field<string>("devid")


                                             } into g
                                             select new dataStruct
                                             {
                                                 文件大小 = g.Sum(p => p.Field<int>("FileSize")),
                                                 在线时长 = g.Sum(p => p.Field<int>("VideLength")),
                                                 GFUploadCnt = g.Sum(p => p.Field<int>("GFUploadCnt")),
                                                 UploadCnt = g.Sum(p => p.Field<int>("UploadCnt")),
                                             }).ToList<dataStruct>();
                                Int64 视频时长 = 0;
                                Int64 文件大小 = 0;
                                int GFUploadCnt = 0;
                                int UploadCnt = 0;
                                int useCnt = 0;
                                foreach (dataStruct item in queryrows)
                                {
                                    视频时长 += item.在线时长;
                                    文件大小 += item.文件大小;
                                    GFUploadCnt += item.GFUploadCnt;
                                    UploadCnt += item.UploadCnt;
                                    useCnt += (item.在线时长 - statusvalue > 0) ? 1 : 0;
                                }
                                dr[2 + h] = useCnt;
                                dr[3 + h] = queryrows.Count - useCnt;
                                dr[4 + h] = Math.Round((double)视频时长 / 3600, 2) ;
                                dr[5 + h] = Math.Round((double)文件大小 / 1048576, 2);
                                dr[6 + h] = (UploadCnt==0)?0:Math.Round((double)GFUploadCnt * 100 / UploadCnt, 2);
                                dr[7 + h] = (queryrows.Count == 0) ? 0 : Math.Round((double)useCnt * 100 / queryrows.Count, 2);

                                drtz[2 + h] = (drtz[2 + h].ToString() == "") ? useCnt : int.Parse(drtz[2 + h].ToString()) + useCnt;
                                drtz[3 + h] = (drtz[3 + h].ToString() == "") ? queryrows.Count - useCnt : int.Parse(drtz[3 + h].ToString()) + queryrows.Count - useCnt;
                                drtz[4 + h] = (drtz[4 + h].ToString() == "") ? Math.Round((double)视频时长 / 3600, 2) : double.Parse(drtz[4 + h].ToString()) + Math.Round((double)视频时长 / 3600, 2);
                                drtz[5 + h] = (drtz[5 + h].ToString() == "") ? Math.Round((double)文件大小 / 1048576, 2) : double.Parse(drtz[5 + h].ToString()) + Math.Round((double)文件大小 / 1048576, 2);

                              if(drtz[6 + h].ToString() == "")
                                {
                                    drtz[6 + h] = GFUploadCnt + "," + UploadCnt;
                                }
                              else
                                {
                                    int tempGFUploadCnt = 0;
                                    int tempUploadCnt = 0;
                                    tempGFUploadCnt = int.Parse(drtz[6 + h].ToString().Split(',')[0]) + GFUploadCnt;
                                    tempUploadCnt = int.Parse(drtz[6 + h].ToString().Split(',')[1])+ UploadCnt;
                                    drtz[6 + h] = (tempGFUploadCnt) + "," + (tempUploadCnt);
                                }


                                if (h == 0)
                                {
                                    dr["1"] = queryrows.Count;
                                    drtz["1"] = (drtz["1"].ToString() == "") ? queryrows.Count : int.Parse(drtz["1"].ToString()) + queryrows.Count;
                                }

                                h += 6;

                            }
                        }
                        catch (Exception e)
                        {

                        }
                        break;

                }


                dtreturns.Rows.Add(dr);
            }

            drtz["0"] = "汇总";//ddtitle;
            switch (type)
            {
                case "1":
                case "2":
                case "3":
                   for (var h=0;h< countTime; h++)
                    {
                        drtz[4+ h*3] = (double.Parse(drtz["1"].ToString())==0)?0: Math.Round(double.Parse(drtz[2 + h*3].ToString())*100/ double.Parse(drtz["1"].ToString()),2);
                    }
                    break;
                case "4":
                case "6":
                    for (var h = 0; h < countTime; h++)
                    {
                        drtz[4 + h * 5] = (double.Parse(drtz["2"].ToString())==0)?0:Math.Round(double.Parse(drtz[3 + h * 5].ToString()) / double.Parse(drtz["2"].ToString()), 2);
                        drtz[6 + h * 5] = (double.Parse(drtz["1"].ToString()) == 0) ? 0 : Math.Round(double.Parse(drtz[3 + h * 5].ToString()) / double.Parse(drtz["1"].ToString()), 2);
                    }
                    break;
                case "5":
                    for (var h = 0; h < countTime; h++)
                    {
                        drtz[6+ h * 6] =(double.Parse(drtz[6 + h * 6].ToString().Split(',')[1])==0)?0:Math.Round(double.Parse(drtz[6 + h * 6].ToString().Split(',')[0]) * 100 / double.Parse(drtz[6 + h * 6].ToString().Split(',')[1]), 2);
                        drtz[7 + h * 6] = (double.Parse(drtz["1"].ToString()) == 0) ? 0 : Math.Round(double.Parse(drtz[2 + h * 6].ToString())*100 / double.Parse(drtz["1"].ToString()), 2);
                    }
                    break;
            }

          


            dtreturns.Rows.Add(drtz);

            insertSheet(dtreturns, sheet, type, typename, reporttype, title);
            if (reporttype != "支队") return;
            foreach (var entityitem in rows)
            {
                if (type != "5" && entityitem["BMDM"].ToString() == "33100000000x") continue;//如果不是执法记录仪，跳出“局机关”单位
                InsertRowdata(sheet, type, typename, entityitem["BMDM"].ToString(), "大队", entityitem["BMMC"].ToString());
            }

        }

        public void insertSheet(DataTable dt, ExcelWorksheet sheet, string type, string typename, string reporttype, string title)
        {
            int mergedint = 0;
            switch (type)
            {
                case "1":
                case "2":
                case "3":
                    mergedint = 1+countTime*3;
                    break;
                case "4":
                case "6":
                    mergedint = 2+ countTime*5;
                    break;
                case "5":
                    mergedint = 1 + countTime * 6;
                    break;
            }
            CellRange range = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows, mergedint);//GetSubrange("A1", "G1");
            range.Value = begintime.Replace("/", "-") + "_" + endtime.Replace("/", "-") + title + typename + "报表";
            range.Merged = true;
            range.Style = Titlestyle();
            sheetrows += 1;
            InsertTitle(sheet, type);//标题添加
            sheet.Rows[0].Cells[0].Style.FillPattern.PatternBackgroundColor = Color.Black;

            for (int h = 0; h < dt.Rows.Count; h++)
            {
                for (int n = 0; n < dt.Columns.Count; n++)
                {
                    sheet.Rows[sheetrows + h].Cells[n].Value = dt.Rows[h][n].ToString();
                    if (dt.Rows[h][n].ToString() != "") sheet.Rows[sheetrows + h].Cells[n].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);

                }

            }
            sheetrows += dt.Rows.Count + 1;
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
            public Int64 在线时长 = 0;
            public Int64 文件大小 = 0;
            public int AlarmType = 0;
            public string DevId = "DevId";
            public int UploadCnt = 0;
            public int GFUploadCnt = 0;
            public int HandleCnt = 0;
            public int CXCnt = 0;
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