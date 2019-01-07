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

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            string type = context.Request.Form["type"];
            string begintime = context.Request.Form["begintime"];
            string endtime = context.Request.Form["endtime"];
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





             allEntitys = SQLHelper.ExecuteRead(CommandType.Text, "SELECT BMDM,SJBM,BMMC,isnull(Sort,0) as Sort,id from [Entity] ", "11");
             devtypes = SQLHelper.ExecuteRead(CommandType.Text, "SELECT TypeName,ID FROM [dbo].[DeviceType] where ID<7  ORDER by Sort ", "11");
             dUser = SQLHelper.ExecuteRead(CommandType.Text, "SELECT en.SJBM,us.BMDM FROM [dbo].[ACL_USER] us left join Entity en on us.BMDM = en.BMDM", "user");
             zfData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT  sum(CONVERT(bigint,[VideLength])) as 视频长度, sum(CONVERT(bigint,[FileSize])) as 文件大小,sum([UploadCnt]) as 上传量,sum([GFUploadCnt]) as 规范上传量,de.BMDM,de.DevId FROM [EveryDayInfo_ZFJLY] al left join Device de on de.DevId = al.DevId where  [Time] >='" + begintime + "' and [Time] <='" + endtime + "'  group by de.DevId,de.BMDM", "Alarm_EveryDayInfo");
             zxscData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT de.BMDM,SUM(value) as value ,de.DevId  FROM [Alarm_EveryDayInfo] al left join Device de on de.DevId = al.DevId where al.AlarmType=1 group by de.DevId,de.BMDM ", "Alarm_EveryDayInfo");
             cllData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT de.BMDM,SUM(value) as  value ,de.DevId  FROM [Alarm_EveryDayInfo] al left join Device de on de.DevId = al.DevId where al.AlarmType=2 group by de.DevId,de.BMDM ", "Alarm_EveryDayInfo");
             cxlData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT de.BMDM,SUM(value) as  value ,de.DevId  FROM [Alarm_EveryDayInfo] al left join Device de on de.DevId = al.DevId where al.AlarmType=5 group by de.DevId,de.BMDM ", "Alarm_EveryDayInfo");







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
                    CellRange range = sheet.Cells.GetSubrange("A1", "G1");
                    range.Value = "台州交警局对讲机报表";
                    range.Merged = true;
                    range.Style = Titlestyle();
                    InsertTitle(sheet, devtypes.Rows[h]["id"].ToString(),1);//标题添加

               





            }


            tmpath = HttpContext.Current.Server.MapPath("upload\\一键报表数据统计.xls");
            excelFile.SaveXls(tmpath);


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

        public void InsertTitle(ExcelWorksheet sheet,string type,int rowindex)
        {
            switch (type)
            {
                case "1":
                case "2":
                case "3":
                    sheet.Rows[rowindex].Cells["A"].Value = "序号";
                    sheet.Rows[rowindex].Cells["B"].Value = "部门";
                    sheet.Rows[rowindex].Cells["C"].Value = "设备配发数(台)";
                    sheet.Rows[rowindex].Cells["D"].Value = "设备使用数量（台）";
                    sheet.Rows[rowindex].Cells["E"].Value = "在线时长总和(小时)";
                    sheet.Rows[rowindex].Cells["F"].Value = "设备使用率";
                    sheet.Rows[rowindex].Cells["G"].Value = "使用率排名";
                     CellRange range = sheet.Cells.GetSubrange("A2", "G2");
                    CellStyle style = new CellStyle();
                    style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                    range.Style = style;
                    //      range.Style.Borders.SetBorders(MultipleBorders.Outside, Color.FromArgb(0, 0, 0), LineStyle.Thin);
                    break;
                case "4":
                case "6":

                    break;
                case "5":

                    break;

            }

        }

        public int InsertRowdata(ExcelWorksheet sheet, string type, int rowindex,string sjbm,string reporttype)
        {
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

            sheetrows = 0;
            string pxstring = "";
            var rows = from p in allEntitys.AsEnumerable()
                       where (p.Field<string>("SJBM") == sjbm)
                       orderby p.Field<int>("Sort") descending
                       select p;
            foreach (var entityitem in rows)
            {
                if (type != "5" && entityitem["BMDM"].ToString() == "33100000000x") continue;//如果不是执法记录仪，跳出“局机关”单位
                DataRow dr = dtreturns.NewRow();
                sheetrows += 1;
                dataindex += 1;
                dr["cloum1"] = dataindex;// entityitem["BMMC"].ToString();  //序号
                dr["cloum2"] = entityitem["BMMC"].ToString();  //部门名称

                var entityids = GetSonID(entityitem["BMDM"].ToString());
                List<string> strList = new List<string>();
                strList.Add(entityitem["BMDM"].ToString());
                if (reporttype == "大队汇总")
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
                                   where strList.ToArray().Contains(p.Field<string>("BMDM"))
                                   select new dataStruct
                                   {
                                       BMDM = p.Field<string>("BMDM"),
                                       在线时长 = p.Field<int>("value"),
                                       DevId = p.Field<string>("DevId")
                                   }).ToList<dataStruct>();
                        dr["cloum3"] = zxrow.Count.ToString();  //配发数
                        int 在线时长=0;
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
                        pxstring = "cloum6";
                        break;
                    case "4":
                    case "6":

                        break;
                    case "5":

                        break;

                }



            }
            DataRow drtz = dtreturns.NewRow();
         
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
                all_pf += (int)item["cloum3"];
                all_use += (int)item["cloum4"];
                all_time += (int)item["cloum5"];
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
            drtz["cloum6"]=(all_pf == 0) ? 0 : Math.Round((double)all_use * 100 / all_pf, 2);//设备使用率
            drtz["cloum7"] = "/";//设备使用率
            dtreturns.Rows.Add(drtz);



            return 0;
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