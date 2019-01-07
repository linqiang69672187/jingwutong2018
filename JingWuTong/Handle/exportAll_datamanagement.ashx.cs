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

        DataTable allEntitys = SQLHelper.ExecuteRead(CommandType.Text, "SELECT BMDM,SJBM,BMMC,isnull(Sort,0) as Sort,id from [Entity] ", "11");
        DataTable devtypes = SQLHelper.ExecuteRead(CommandType.Text, "SELECT TypeName,ID FROM [dbo].[DeviceType] where ID<7  ORDER by Sort ", "11");
        DataTable dUser = SQLHelper.ExecuteRead(CommandType.Text, "SELECT en.SJBM,us.BMDM FROM [dbo].[ACL_USER] us left join Entity en on us.BMDM = en.BMDM", "user");
        DataTable zfData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT  sum(CONVERT(bigint,[VideLength])) as 视频长度, sum(CONVERT(bigint,[FileSize])) as 文件大小,sum([UploadCnt]) as 上传量,sum([GFUploadCnt]) as 规范上传量,de.BMDM,de.DevId FROM [EveryDayInfo_ZFJLY] al left join Device de on de.DevId = al.DevId where  [Time] >='" + begintime + "' and [Time] <='" + endtime + "'  group by de.DevId,de.BMDM", "Alarm_EveryDayInfo");
        DataTable zxscData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT de.BMDM,SUM(value) as value ,de.DevId  FROM [Alarm_EveryDayInfo] al left join Device de on de.DevId = al.DevId where al.AlarmType=1 group by de.DevId,de.BMDM ", "Alarm_EveryDayInfo");
        DataTable cllData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT de.BMDM,SUM(value) as  value ,de.DevId  FROM [Alarm_EveryDayInfo] al left join Device de on de.DevId = al.DevId where al.AlarmType=2 group by de.DevId,de.BMDM ", "Alarm_EveryDayInfo");
        DataTable cxlData = SQLHelper.ExecuteRead(CommandType.Text, "SELECT de.BMDM,SUM(value) as  value ,de.DevId  FROM [Alarm_EveryDayInfo] al left join Device de on de.DevId = al.DevId where al.AlarmType=5 group by de.DevId,de.BMDM ", "Alarm_EveryDayInfo");

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

  

     








            int days = Convert.ToInt16(context.Request.Form["dates"]);
            int statusvalue = 10;  //正常参考值
            int zxstatusvalue = 30;//在线参考值
            int sheetrows = 0;
            int dataindex = 0;

            statusvalue = days * usedvalue;//超过10分钟算使用
            zxstatusvalue = days * onlinevalue;//在线参考值


            ExcelFile excelFile = new ExcelFile();
            var tmpath = "";
            tmpath = HttpContext.Current.Server.MapPath("templet\\0.xls");
            excelFile.LoadXls(tmpath);
            //所有大队

            for (int h = 0; h < devtypes.Rows.Count; h++)
            {
                sheetrows = 0;
                var rows = from p in allEntitys.AsEnumerable()
                           where (p.Field<string>("SJBM")== "331000000000")
                           orderby p.Field<int>("Sort") descending
                           select p;
                ExcelWorksheet sheet = excelFile.Worksheets[devtypes.Rows[h]["TypeName"].ToString()];
                    CellRange range = sheet.Cells.GetSubrange("A1", "G1");
                    range.Value = "台州交警局对讲机报表";
                    range.Merged = true;
                    range.Style = Titlestyle();
                    InsertTitle(sheet, devtypes.Rows[h]["id"].ToString(),1);//标题添加

                foreach (var entityitem in rows)
                        {
                            if (devtypes.Rows[h]["TypeName"].ToString() != "执法记录仪"&& entityitem["BMDM"].ToString()== "33100000000x") continue;
                             sheetrows += 1;
                             dataindex += 1;
                             sheet.Rows[sheetrows + 1].Cells["A"].Value = dataindex;// entityitem["BMMC"].ToString();
                             sheet.Rows[sheetrows + 1].Cells["B"].Value = entityitem["BMMC"].ToString();
                      



                        }





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

        public int InsertRowdata(ExcelWorksheet sheet, string type, int rowindex,DataTable dt,string entityid,string reporttype)
        {

            var entityids = GetSonID(entityid);
            List<string> strList = new List<string>();
            strList.Add(entityid);
            if (reporttype == "大队汇总") { 
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


                    break;
                case "4":
                case "6":

                    break;
                case "5":

                    break;

            }

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