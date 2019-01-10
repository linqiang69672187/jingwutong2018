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


        List<dataStruct> tmpList = new List<dataStruct>();

        DataTable allEntitys = null;
        DataTable devtypes = null;
        DataTable dUser = null;
        DataTable zfData = null;
        DataTable zxscData = null;
        DataTable cllData = null;
        DataTable cxlData = null;


        int statusvalue = 0;  //正常参考值
        int zxstatusvalue = 0;//在线参考值

        int sheetrows = 0;
        int dataindex = 0;
        string begintime = "";
        string endtime = "";

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
                sreachcondi = " (de.[DevId] like '%" + search + "%' or us.[XM] like '%" + search + "%' or us.[JYBH] like '%" + search + "%' ) and ";
            }





             allEntitys = SQLHelper.ExecuteRead(CommandType.Text, "SELECT BMDM,SJBM,BMMC,isnull(Sort,0) as Sort,id from [Entity] ", "11");
             devtypes = SQLHelper.ExecuteRead(CommandType.Text, "SELECT TypeName,ID FROM [dbo].[DeviceType] where ID<7  ORDER by Sort ", "11");
             dUser = SQLHelper.ExecuteRead(CommandType.Text, "SELECT en.SJBM,us.BMDM FROM [dbo].[ACL_USER] us left join Device de on de.JYBH = us.JYBH left join Entity en on de.BMDM = en.BMDM where " + sreachcondi + " us.[JYBH] <>''", "user");
             zfData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT  sum(CONVERT(bigint,[VideLength])) as 视频长度, sum(CONVERT(bigint,[FileSize])) as 文件大小,sum([UploadCnt]) as 上传量,sum([GFUploadCnt]) as 规范上传量,de.BMDM,de.DevId FROM [EveryDayInfo_ZFJLY] al left join Device de on de.DevId = al.DevId left join ACL_USER as us on de.JYBH = us.JYBH  where " + sreachcondi + "   [Time] >='" + begintime + "' and [Time] <='" + endtime + "'  group by de.DevId,de.BMDM", "Alarm_EveryDayInfo");
             zxscData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT de.BMDM,SUM(value) as value ,de.DevId,de.devtype  FROM [Alarm_EveryDayInfo] al left join Device de on de.DevId = al.DevId left join ACL_USER as us on de.JYBH = us.JYBH  where  " + sreachcondi + "  [AlarmDay ] >='" + begintime + "' and [AlarmDay ] <='" + endtime + "'  and al.AlarmType=1 group by de.DevId,de.BMDM,de.devtype ", "Alarm_EveryDayInfo");
             cllData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT de.BMDM,SUM(value) as  value ,de.DevId ,de.devtype FROM [Alarm_EveryDayInfo] al left join Device de on de.DevId = al.DevId left join ACL_USER as us on de.JYBH = us.JYBH  where " + sreachcondi + "  [AlarmDay ] >='" + begintime + "' and [AlarmDay ] <='" + endtime + "'  and al.AlarmType=2 group by de.DevId,de.BMDM,de.devtype", "Alarm_EveryDayInfo");
             cxlData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT de.BMDM,SUM(value) as  value ,de.DevId ,de.devtype FROM [Alarm_EveryDayInfo] al left join Device de on de.DevId = al.DevId left join ACL_USER as us on de.JYBH = us.JYBH  where " + sreachcondi + "  [AlarmDay ] >='" + begintime + "' and [AlarmDay ] <='" + endtime + "'  and al.AlarmType=5 group by de.DevId,de.BMDM,de.devtype", "Alarm_EveryDayInfo");







            int days = Convert.ToInt16(context.Request.Form["dates"]);
           
          

            statusvalue = days * usedvalue;//超过10分钟算使用
            zxstatusvalue = days * onlinevalue;//在线参考值


            ExcelFile excelFile = new ExcelFile();
            var tmpath = "";
            tmpath = HttpContext.Current.Server.MapPath("templet\\0.xls");
            excelFile.LoadXls(tmpath);
            //所有大队

            for (int h = 0; h < devtypes.Rows.Count; h++)
            {
              
                ExcelWorksheet sheet = excelFile.Worksheets[devtypes.Rows[h]["TypeName"].ToString()];
                sheetrows = 0;
                InsertRowdata(sheet, devtypes.Rows[h]["id"].ToString(), devtypes.Rows[h]["TypeName"].ToString(), "331000000000", "支队","台州交警局");




            }


            tmpath = HttpContext.Current.Server.MapPath("upload\\" + begintime.Replace("/", "-") + "_" + endtime.Replace("/", "-") + "日报表.xls");
            excelFile.SaveXls(tmpath);
            context.Response.Redirect(tmpath);

            //string reTitle = ExportExcel(dtreturns, type, begintime, endtime, ssdd, sszd);

        }

        public CellStyle Titlestyle()
        {
            CellStyle style = new CellStyle();
            //设置水平对齐模式
            style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            //设置垂直对齐模式
            style.VerticalAlignment = VerticalAlignmentStyle.Center;
            //设置字体
            style.Font.Size =12*20; //PT=20
            style.Font.Weight = ExcelFont.BoldWeight;
                 
            //style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
            //  style.Font.Color = Color.Blue;
            return style;
        }

        public void InsertTitle(ExcelWorksheet sheet,string type)
        {
            CellRange range;
            CellStyle style;
            switch (type)
            {
                case "1":
                case "2":
                case "3":
                    sheet.Rows[sheetrows].Cells["A"].Value = "序号";
                    sheet.Rows[sheetrows].Cells["B"].Value = "部门";
                    sheet.Rows[sheetrows].Cells["C"].Value = "设备配发数(台)";
                    sheet.Rows[sheetrows].Cells["D"].Value = "设备使用数量（台）";
                    sheet.Rows[sheetrows].Cells["E"].Value = "在线时长总和(小时)";
                    sheet.Rows[sheetrows].Cells["F"].Value = "设备使用率";
                    sheet.Rows[sheetrows].Cells["G"].Value = "使用率排名";
                     range = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows, 6);
                     style = new CellStyle();
                    style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                    range.Style = style;
                    //      range.Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                    break;
                case "4":
                    sheet.Rows[sheetrows].Cells["A"].Value = "序号";
                    sheet.Rows[sheetrows].Cells["B"].Value = "部门";
                    sheet.Rows[sheetrows].Cells["C"].Value = "设备配发数(台)";
                    sheet.Rows[sheetrows].Cells["D"].Value = "警员数";
                    sheet.Rows[sheetrows].Cells["E"].Value = "警务通处罚数";
                    sheet.Rows[sheetrows].Cells["F"].Value = "人均处罚量";
                    sheet.Rows[sheetrows].Cells["G"].Value = "查询量";
                    sheet.Rows[sheetrows].Cells["H"].Value = "设备平均处罚量";
                    sheet.Rows[sheetrows].Cells["I"].Value = "设备平均处罚量排名";
                    sheet.Rows[sheetrows].Cells["J"].Value = "无处罚数的警务通（台）";
                    range = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows,9);
                    style = new CellStyle();
                    style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                    range.Style = style;
                    break;
                case "6":
                    sheet.Rows[sheetrows].Cells["A"].Value = "序号";
                    sheet.Rows[sheetrows].Cells["B"].Value = "部门";
                    sheet.Rows[sheetrows].Cells["C"].Value = "设备配发数(台)";
                    sheet.Rows[sheetrows].Cells["D"].Value = "辅警数";
                    sheet.Rows[sheetrows].Cells["E"].Value = "违停采集（例）";
                    sheet.Rows[sheetrows].Cells["F"].Value = "人均处罚量";
                    sheet.Rows[sheetrows].Cells["G"].Value = "查询量";
                    sheet.Rows[sheetrows].Cells["H"].Value = "设备平均处罚量";
                    sheet.Rows[sheetrows].Cells["I"].Value = "设备平均处罚量排名";
                    sheet.Rows[sheetrows].Cells["J"].Value = "本月无采集违停设备（台）";
                    range = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows, 9);
                    style = new CellStyle();
                    style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                    range.Style = style;
                    break;
                case "5":
                    sheet.Rows[sheetrows].Cells["A"].Value = "序号";
                    sheet.Rows[sheetrows].Cells["B"].Value = "部门";
                    sheet.Rows[sheetrows].Cells["C"].Value = "设备配发数(台)";
                    sheet.Rows[sheetrows].Cells["D"].Value = "设备使用数量（台）";
                    sheet.Rows[sheetrows].Cells["E"].Value = "设备未使用数量（台）";
                    sheet.Rows[sheetrows].Cells["F"].Value = "视频时长总和(小时)";
                    sheet.Rows[sheetrows].Cells["G"].Value = "视频大小(GB)";
                    sheet.Rows[sheetrows].Cells["H"].Value = "设备使用率";
                    sheet.Rows[sheetrows].Cells["I"].Value = "使用率排名";
                     range = sheet.Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows, 8);
                     style = new CellStyle();
                    style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                    range.Style = style;
                    break;

            }
            sheetrows += 1;
        }

        public void InsertRowdata(ExcelWorksheet sheet, string type,string typename,string sjbm,string reporttype,string title)
        {
            DataTable dtreturns = new DataTable(); //返回数据表
            dtreturns.Columns.Add("cloum1");
            dtreturns.Columns.Add("cloum2");
            dtreturns.Columns.Add("cloum3");
            dtreturns.Columns.Add("cloum4");
            dtreturns.Columns.Add("cloum5");
            dtreturns.Columns.Add("cloum6", typeof(double));
            dtreturns.Columns.Add("cloum7");
            dtreturns.Columns.Add("cloum8", typeof(double));
            dtreturns.Columns.Add("cloum9");
            dtreturns.Columns.Add("cloum10");
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
                       where (p.Field<string>("SJBM") == sjbm|| p.Field<string>("BMDM") == sjbm)
                       orderby p.Field<int>("Sort") descending
                       select p;
            }
            foreach (var entityitem in rows)
            {
                if (type != "5" && entityitem["BMDM"].ToString() == "33100000000x") continue;//如果不是执法记录仪，跳出“局机关”单位
                DataRow dr = dtreturns.NewRow();
                dataindex += 1;
                dr["cloum1"] = dataindex;// entityitem["BMMC"].ToString();  //序号
                dr["cloum2"] = entityitem["BMMC"].ToString();  //部门名称

                var entityids = GetSonID(entityitem["BMDM"].ToString());
                List<string> strList = new List<string>();
                strList.Add(entityitem["BMDM"].ToString());
                if (!(reporttype != "支队"&& entityitem["SJBM"].ToString()=="331000000000"))  //非支队报表下的大队单位，只显示本级
                {
                    foreach (entityStruct item in entityids)
                    {
                        strList.Add(item.BMDM);
                    }
                }
                switch (type)
                {
                    case "1":
                    case "2":
                    case "3":
                        var zxrow = (from p in zxscData.AsEnumerable()
                                   where strList.ToArray().Contains(p.Field<string>("BMDM")) && p.Field<int>("devtype").ToString()==type
                                     select new dataStruct
                                   {
                                       BMDM = p.Field<string>("BMDM"),
                                       在线时长 = p.Field<int>("value"),
                                       DevId = p.Field<string>("DevId")
                                   }).ToList<dataStruct>();
                        dr["cloum3"] = zxrow.Count.ToString();  //配发数
                        Int64 在线时长 =0;
                        int 设备使用台数=0;
                        foreach (var row in zxrow)
                        {
                            在线时长 += row.在线时长;  
                            设备使用台数+=((Convert.ToInt32(row.在线时长) - statusvalue) > 0) ? 1 : 0;
                        }
                        dr["cloum4"] = 设备使用台数;//设备使用数量
                        dr["cloum5"] =  Math.Round((double)在线时长 / 3600, 2);//设备使用数量
                        dr["cloum6"] = (zxrow.Count==0)?0: Math.Round((double)设备使用台数* 100 / zxrow.Count, 2);//设备使用率
                        dtreturns.Rows.Add(dr);
                        break;
                    case "4":
                    case "6":
                        try
                        {
                            var cxlrow = (from p in cxlData.AsEnumerable()
                                         where strList.ToArray().Contains(p.Field<string>("BMDM"))
                                         select new dataStruct
                                         {
                                             BMDM = p.Field<string>("BMDM"),
                                             CXCnt = p.Field<int>("value"),
                                             DevId = p.Field<string>("DevId")
                                         }).ToList<dataStruct>();
                            var cllrow = (from p in cllData.AsEnumerable()
                                          where strList.ToArray().Contains(p.Field<string>("BMDM"))
                                          select new dataStruct
                                          {
                                              BMDM = p.Field<string>("BMDM"),
                                              HandleCnt = p.Field<int>("value"),
                                              DevId = p.Field<string>("DevId")
                                          }).ToList<dataStruct>();
                            var userrow = (from p in dUser.AsEnumerable()
                                          where strList.ToArray().Contains(p.Field<string>("BMDM"))
                                          select p);
                            dr["cloum3"] = cllrow.Count().ToString();  //配发数
                            dr["cloum4"] = userrow.Count().ToString();
                            int 处罚数 = 0;
                            int 查询量 = 0;
                            int 无处罚量设备 = 0;
                            foreach (var row in cxlrow)
                            {
                                查询量 += row.CXCnt;
                            }
                            foreach (var row in cllrow)
                            {
                                处罚数 += row.HandleCnt;
                                无处罚量设备 += 1;
                            }
                            dr["cloum5"] = 处罚数;//设备使用数量
                            dr["cloum6"] = (userrow.Count() == 0) ? 0 : Math.Round((double)处罚数 * 100 / userrow.Count(), 2);//人均处罚量
                            dr["cloum7"] = 查询量;//查询量
                            dr["cloum8"] = (cllrow.Count() == 0) ? 0 : Math.Round((double)处罚数 * 100 / cllrow.Count(), 2);//查询量
                            dr["cloum10"] = 无处罚量设备;

                            dtreturns.Rows.Add(dr);
                        }
                        catch (Exception e)
                        {

                        }
                        break;
                    case "5":
                        try { 
                        var zfrow = (from p in zfData.AsEnumerable()
                                     where strList.ToArray().Contains(p.Field<string>("BMDM"))
                                     select new dataStruct
                                     {
                                         BMDM = p.Field<string>("BMDM"),
                                         在线时长 = p.Field<Int64>("视频长度"),
                                         文件大小 = p.Field<Int64>("文件大小"),
                                         DevId = p.Field<string>("DevId")
                                     }).ToList<dataStruct>();
                        dr["cloum3"] = zfrow.Count.ToString();  //配发数
                         Int64 视频时长 = 0;
                        int 执法记录仪设备使用台数 = 0;
                        Int64 文件大小 = 0;
                        foreach (var row in zfrow)
                        {
                            视频时长 += row.在线时长;
                            文件大小 += row.文件大小;
                            执法记录仪设备使用台数 += ((Convert.ToInt32(row.在线时长) - statusvalue) > 0) ? 1 : 0;
                        }
                        dr["cloum4"] = 执法记录仪设备使用台数;//设备使用数量

                        dr["cloum5"] = int.Parse(zfrow.Count.ToString()) - 执法记录仪设备使用台数;
                        dr["cloum6"] = Math.Round((double)视频时长 / 3600, 2);
                        dr["cloum7"] = Math.Round((double)文件大小 / 1048576, 2);
                        dr["cloum8"] = (zfrow.Count == 0) ? 0 : Math.Round((double)执法记录仪设备使用台数 * 100 / zfrow.Count, 2);//设备使用率
                        dtreturns.Rows.Add(dr);
                        }
                        catch (Exception e) {

                        }
                        break;

                }



            }
            DataRow drtz = dtreturns.NewRow();
            switch (type)
            {
                case "1":
                case "2":
                case "3":
                    pxstring = "cloum6";
                    int all_pf = 0;
                    int all_use = 0;
                    double all_time = 0.0;
                    int orderno = 1;
                    var query = (from p in dtreturns.AsEnumerable()
                                 orderby p.Field<double>(pxstring) descending
                                 select p) as IEnumerable<DataRow>;
                    double temsyl = 0.0;
                    int temorder = 1;
                    foreach (var item in query)
                    {
                        all_pf += int.Parse(item["cloum3"].ToString());
                        all_use += int.Parse(item["cloum4"].ToString());
                        all_time += double.Parse(item["cloum5"].ToString());
                        if (temsyl == double.Parse(item[pxstring].ToString()))
                        {
                            item["cloum7"] = temorder;
                        }
                        else
                        {
                            item["cloum7"] = orderno;

                            temsyl = double.Parse((item[pxstring].ToString()));
                            temorder = orderno;
                        }
                        orderno += 1;
                    }

                    drtz["cloum1"] = dtreturns.Rows.Count + 1;
                    drtz["cloum2"] = "合计";//ddtitle;
                    drtz["cloum3"] = all_pf;
                    drtz["cloum4"] = all_use;
                    drtz["cloum5"] = all_time;
                    drtz["cloum6"] = (all_pf == 0) ? 0 : Math.Round((double)all_use * 100 / all_pf, 2);//设备使用率
                    drtz["cloum7"] = "/";//设备使用率
                    break;
                case "4":
                case "6":
                    try
                    {
                        pxstring = "cloum8";
                        int jwt_all_pf = 0;
                        int jwt_all_user = 0;
                        int jwt_all_cfs = 0;
                        int jwt_all_cxl = 0;
                        int jwt_all_wchf = 0;
                        int jwt_orderno = 1;
                        double jwt_temsyl = 0.0;
                        int jwt_temorder = 1;

                        var jwtquery = (from p in dtreturns.AsEnumerable()
                                       orderby p.Field<double>(pxstring) descending
                                       select p) as IEnumerable<DataRow>;

                        foreach (var item in jwtquery)
                        {
                            jwt_all_pf += int.Parse(item["cloum3"].ToString());
                            jwt_all_user += int.Parse(item["cloum4"].ToString());
                            jwt_all_cfs += int.Parse(item["cloum5"].ToString());
                            jwt_all_cxl =+int.Parse(item["cloum7"].ToString());
                            jwt_all_wchf += int.Parse(item["cloum10"].ToString());
                            if (jwt_temsyl == double.Parse(item[pxstring].ToString()))
                            {
                                item["cloum9"] = jwt_temorder;
                            }
                            else
                            {
                                item["cloum9"] = jwt_orderno;

                                jwt_temsyl = double.Parse((item[pxstring].ToString()));
                                jwt_temorder = jwt_orderno;
                            }
                            jwt_orderno += 1;
                        }

                        drtz["cloum1"] = dtreturns.Rows.Count + 1;
                        drtz["cloum2"] = "合计";//ddtitle;
                        drtz["cloum3"] = jwt_all_pf;
                        drtz["cloum4"] = jwt_all_user;
                        drtz["cloum5"] = jwt_all_cfs;
                        drtz["cloum6"] = (jwt_all_user == 0) ? 0 : Math.Round((double)jwt_all_cfs * 100 / jwt_all_user, 2); ;
                        drtz["cloum7"] = jwt_all_cxl;
                        drtz["cloum8"] = (jwt_all_pf == 0) ? 0 : Math.Round((double)jwt_all_cfs * 100 / jwt_all_pf, 2);//设备使用率
                        drtz["cloum9"] = "/";
                        drtz["cloum10"] = jwt_all_wchf;
                    }
                    catch (Exception e)
                    {

                    }
                    break;
                case "5":
                    try {
                    pxstring = "cloum8";
                    int zf_all_pf = 0;
                    int zf_all_use = 0;
                    int zf_all_unuse = 0;
                    double all_视频时长 = 0.0;
                    double all_文件大小 = 0.0;
                    int zf_orderno = 1;
                    double zf_temsyl = 0.0;
                    int zf_temorder = 1;

                    var zfquery = (from p in dtreturns.AsEnumerable()
                                 orderby p.Field<double>(pxstring) descending
                                 select p) as IEnumerable<DataRow>;
                
                    foreach (var item in zfquery)
                    {
                            zf_all_pf += int.Parse(item["cloum3"].ToString());
                            zf_all_use += int.Parse(item["cloum4"].ToString());
                            zf_all_unuse += int.Parse(item["cloum5"].ToString());
                            all_视频时长 += double.Parse(item["cloum6"].ToString());
                            all_文件大小 += double.Parse(item["cloum7"].ToString());
                            if (zf_temsyl == double.Parse(item[pxstring].ToString()))
                            {
                                item["cloum9"] = zf_temorder;
                            }
                            else
                            {
                                item["cloum9"] = zf_orderno;

                                zf_temsyl = double.Parse((item[pxstring].ToString()));
                                zf_temorder = zf_orderno;
                            }
                            zf_orderno += 1;
                        }

                    drtz["cloum1"] = dtreturns.Rows.Count + 1;
                    drtz["cloum2"] = "合计";//ddtitle;
                    drtz["cloum3"] = zf_all_pf;
                    drtz["cloum4"] = zf_all_use;
                    drtz["cloum5"] = zf_all_unuse;
                    drtz["cloum6"] = all_视频时长;
                    drtz["cloum7"] = all_文件大小;
                    drtz["cloum8"] = (zf_all_pf == 0) ? 0 : Math.Round((double)zf_all_use * 100 / zf_all_pf, 2);//设备使用率
                    drtz["cloum9"] = "/";//设备使用率
                    }
                    catch (Exception e) {

                    }
                    break;

            }
            dtreturns.Rows.Add(drtz);
            insertSheet(dtreturns,  sheet,type,typename,reporttype,title);
            if (reporttype != "支队") return;
            foreach (var entityitem in rows)
            {
                if (type != "5" && entityitem["BMDM"].ToString() == "33100000000x") continue;//如果不是执法记录仪，跳出“局机关”单位
                InsertRowdata(sheet, type,typename, entityitem["BMDM"].ToString(), "大队",entityitem["BMMC"].ToString());
            }
                
        }

        public void insertSheet(DataTable dt,  ExcelWorksheet sheet,string type,string typename, string reporttype,string title)
        {
           int  mergedint = 0;
            switch (type)
            {
                case "1":
                case "2":
                case "3":
                    mergedint = 6;
                    break;
                case "4":
                case "6":
                    mergedint = 9;
                    break;
                case "5":
                    mergedint = 8;
                    break;
            }
            CellRange range = sheet.Rows[sheetrows].Cells.GetSubrangeAbsolute(sheetrows, 0, sheetrows, mergedint);//GetSubrange("A1", "G1");
            range.Value = begintime.Replace("/", "-") + "_"+ endtime.Replace("/", "-") + title +typename+ "报表";
            range.Merged = true;
            range.Style = Titlestyle();
            sheetrows += 1;
            InsertTitle(sheet, type);//标题添加

       
            for (int h = 0; h < dt.Rows.Count; h++)
            {
                for (int n = 0; n< dt.Columns.Count; n++)
                {
                    sheet.Rows[sheetrows + h].Cells[n].Value = dt.Rows[h][n].ToString();
                    if(dt.Rows[h][n].ToString()!="") sheet.Rows[sheetrows + h].Cells[n].Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);

                }

            }
            sheetrows += dt.Rows.Count+1;
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