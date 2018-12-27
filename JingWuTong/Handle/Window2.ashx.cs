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
    /// Window2 的摘要说明
    /// </summary>
    public class Window2 : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            string type = context.Request.Form["type"];
            string begintime = context.Request.Form["begintime"];
            string endtime = context.Request.Form["endtime"];
            string hbbegintime = context.Request.Form["hbbegintime"];
            string hbendtime = context.Request.Form["hbendtime"];
            //string ssdd = context.Request.Form["ssdd"];
            string sszd = context.Request.Form["sszd"];
            string requesttype = context.Request.Form["requesttype"];
            //string search = context.Request.Form["search"];
            string sreachcondi = "";

            string URL = context.Request.Form["URL"];

            //int onlinevalue = int.Parse(context.Request.Form["onlinevalue"])*60;
            //int usedvalue= int.Parse(context.Request.Form["usedvalue"]) * 60;


            //if (search != "")
            //{
            //    sreachcondi = " (de.[DevId] like '%" + search + "%' or us.[XM] like '%" + search + "%' or us.[JYBH] like '%" + search + "%' ) and ";
            //}

            string tmpDevid = "";
            int tmpRows = 0;
            DataTable dtEntity = null;  //单位信息表
            //typetext: typetext, ssddtext: ssddtext, sszdtext: sszdtext,
            // endtime = "2017/9/14"; //测试使用
            string title = "";
            if (context.Request.Form["ssddtext"] == "全部")
            {
                title = "台州交警局";
            }
            else
            {
                title = context.Request.Form["ssddtext"];
            }

            if (context.Request.Form["sszdtext"] != "全部")
            {
                title += context.Request.Form["sszdtext"];
            }



            StringBuilder sqltext = new StringBuilder();

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
            dtreturns.Columns.Add("cloum11");
            dtreturns.Columns.Add("cloum13", typeof(int));

            int days = Convert.ToInt16(context.Request.Form["dates"]);
            int statusvalue = 10;  //正常参考值
            int zxstatusvalue = 30;//在线参考值
            int devicescount = 0;  //汇总设备总数
            double zxsc = 0.0;  //汇总在线时长
            double spdx = 0.0;  //汇总视频大小
            Int64 cxl = 0;  //汇总查询量
            int wcxl = 0;   //无查询量设备数量
            int wcfl = 0;   //无处罚量设备数量
            int wsysb = 0;  //无使用设备数量
            int zxsb = 0; //在线设备






            string bmdm = ""; //汇总的部门代码
            int allstatu_device = 0;  //汇总使用率不为空数量
            string ddtitle;//大队标题

            //----------



            //statusvalue = days * usedvalue;//超过10分钟算使用
            //zxstatusvalue = days * onlinevalue;//在线参考值
            DataTable Alarm_EveryDayInfo = null; //每日告警
            DataTable dUser = null;


            bmdm = sszd;
            ddtitle = context.Request.Form["sszdtext"];
            switch (type)
            {
                case "5":
                    Alarm_EveryDayInfo = SQLHelper.ExecuteRead(CommandType.Text, "SELECT en.BMQC,en.BMDM ,en.BMDM as [ParentID],us.XM as [Contacts],de.[DevId],ala.在线时长,ala.[AlarmType],us.JYBH,us.SFZMHM from (SELECT [DevId],sum([VideLength]) as 在线时长,1 as AlarmType from [EveryDayInfo_ZFJLY]   where  [Time] >='" + begintime + "' and [Time] <='" + endtime + "'   group by [DevId] ) as ala left join [Device] as de on de.[DevId] = ala.[DevId] left join [Entity] as en on en.[BMDM] = de.[BMDM] left join ACL_USER as us on de.JYBH = us.JYBH  where  de.[DevType]=" + type + " and  en.BMDM='" + sszd + "'  and us.XM  IS NOT NULL ", "Alarm_EveryDayInfo");
                    break;
                default:
                    Alarm_EveryDayInfo = SQLHelper.ExecuteRead(CommandType.Text, "SELECT en.BMQC,en.BMDM ,en.BMDM as [ParentID],us.XM as [Contacts],de.[DevId],ala.在线时长,ala.[AlarmType],us.JYBH,us.SFZMHM from (SELECT [DevId],[AlarmType],sum([Value]) as 在线时长 from [Alarm_EveryDayInfo]   where [AlarmType] <> 6 and  [AlarmDay ] >='" + begintime + "' and [AlarmDay ] <='" + endtime + "'   group by [DevId],[AlarmType] ) as ala left join [Device] as de on de.[DevId] = ala.[DevId] left join [Entity] as en on en.[BMDM] = de.[BMDM] left join ACL_USER as us on de.JYBH = us.JYBH  where  de.[DevType]=" + type + " and en.BMDM='" + sszd + "' and us.XM  IS NOT NULL ", "Alarm_EveryDayInfo");
                    break;
            }


            //dtEntity = SQLHelper.ExecuteRead(CommandType.Text, "  SELECT BMDM as ID,BMQC as Name,SJBM as ParentID,BMJB AS Depth,Sort from [Entity] a       where ( SJBM='" + sszd + "') and [BMQC] IS NOT NULL AND BMQC <> '' ", "2");

            dUser = SQLHelper.ExecuteRead(CommandType.Text, "SELECT en.SJBM,us.BMDM FROM [dbo].[ACL_USER] us left join Entity en on us.BMDM = en.BMDM  ", "user");



            for (int i1 = 0; i1 < Alarm_EveryDayInfo.Rows.Count; i1++)
            {

                DataRow dr = dtreturns.NewRow();

                dr["cloum1"] = (i1 + 1).ToString();
                dr["cloum2"] = Alarm_EveryDayInfo.Rows[i1]["BMQC"].ToString();
                dr["cloum13"] = (i1 + 1);


                Int64 在线时长 = 0;
                Int64 视频大小 = 0;
                Int64 处理量 = 0;
                Int64 文件大小 = 0;
                Int64 查询量 = 0;
                int 无查询量 = 0;
                int 无处罚量 = 0;
                int 未使用 = 0;
                int usercount = 0;
                int 在线 = 0;
                int status = 0;//设备使用正常、周1次，月4次，季度12次
                //-----------

                string s_Xm = "";
                string s_Devid = "";
                string s_JYbh = "";
                string s_SFZMHM = "";

                //var rows = from p in Alarm_EveryDayInfo.AsEnumerable()
                //           orderby p.Field<string>("DevId")
                //           select p;
                //获得设备数量，及正常使用设备
                tmpRows = 0;
                //foreach (var item in rows)
                //{
                if (Alarm_EveryDayInfo.Rows[i1]["在线时长"] is DBNull) { }
                    else
                    {
                        s_Xm = Alarm_EveryDayInfo.Rows[i1]["Contacts"].ToString();//警员姓名
                        s_JYbh = Alarm_EveryDayInfo.Rows[i1]["JYBH"].ToString();//警员编号
                        s_Devid = Alarm_EveryDayInfo.Rows[i1]["DevId"].ToString();//设备编号
                        s_SFZMHM = Alarm_EveryDayInfo.Rows[i1]["SFZMHM"].ToString();//身份证号
                        switch (Alarm_EveryDayInfo.Rows[i1]["AlarmType"].ToString())
                        {
                            case "1":
                                在线时长 += Convert.ToInt32(Alarm_EveryDayInfo.Rows[i1]["在线时长"]);
                                未使用 += ((Convert.ToInt32(Alarm_EveryDayInfo.Rows[i1]["在线时长"]) - statusvalue) <= 0) ? 1 : 0;
                                在线 += ((Convert.ToInt32(Alarm_EveryDayInfo.Rows[i1]["在线时长"]) - zxstatusvalue) > 0) ? 1 : 0;
                                break;
                            case "2":
                                处理量 += Convert.ToInt32(Alarm_EveryDayInfo.Rows[i1]["在线时长"]);
                                无处罚量 += (Convert.ToInt32(Alarm_EveryDayInfo.Rows[i1]["在线时长"]) == 0) ? 1 : 0;
                                break;
                            case "3":
                                文件大小 += Convert.ToInt32(Alarm_EveryDayInfo.Rows[i1]["在线时长"]);
                                break;
                            case "4":
                                视频大小 += Convert.ToInt32(Alarm_EveryDayInfo.Rows[i1]["在线时长"]);
                                break;
                            case "5":
                                查询量 += Convert.ToInt32(Alarm_EveryDayInfo.Rows[i1]["在线时长"]);
                                无查询量 += (Convert.ToInt32(Alarm_EveryDayInfo.Rows[i1]["在线时长"]) == 0) ? 1 : 0;
                                break;
                        }
                        if (Alarm_EveryDayInfo.Rows[i1]["DevId"].ToString() != tmpDevid)
                        {
                            tmpRows += 1;  //新设备ID不重复
                            tmpDevid = Alarm_EveryDayInfo.Rows[i1]["DevId"].ToString();
                            status += (Convert.ToInt32(Alarm_EveryDayInfo.Rows[i1]["在线时长"]) - statusvalue > 0) ? 1 : 0;
                            allstatu_device += (Convert.ToInt32(Alarm_EveryDayInfo.Rows[i1]["在线时长"]) - statusvalue > 0) ? 1 : 0;
                        }

                    }


                //}

                var userrows = from p in dUser.AsEnumerable()
                               where (p.Field<string>("SJBM") == Alarm_EveryDayInfo.Rows[i1]["BMDM"].ToString() || p.Field<string>("BMDM") == Alarm_EveryDayInfo.Rows[i1]["BMDM"].ToString())
                               select p;
                usercount = userrows.Count();

                int countdevices = tmpRows;
                double deviceuse = Math.Round((double)status * 100 / (double)countdevices, 2);

                //dr["cloum3"] = countdevices;//配发台数
                //devicescount += countdevices;
                //Allotment += countdevices;//汇总配发台数

                switch (type)
                {
                    case "4"://警务通
                        dr["cloum3"] = s_Xm;
                        dr["cloum4"] = s_JYbh;
                        dr["cloum5"] = s_Devid;
                        dr["cloum6"] = 处理量;//警务通处罚量
                        dr["cloum7"] = 查询量;
                        break;

                    case "6"://辅警通
                        dr["cloum3"] = s_Xm;
                        dr["cloum4"] = s_JYbh;
                        dr["cloum5"] = s_SFZMHM;
                        dr["cloum6"] = 处理量;//违停从采集量
                        dr["cloum7"] = 查询量;
                        break;
                    case "5":   //执法记录仪
                        dr["cloum3"] = s_Xm;
                        dr["cloum4"] = s_JYbh;
                        dr["cloum5"] =s_Devid;

                        dr["cloum6"] = 视频大小;
                        dr["cloum7"] = 文件大小;//视频大小（GB）
                    
                        break;
                    case "1":
                    case "2":
                    case "3":
                    case "7":
                        dr["cloum3"] = s_Xm;
                        dr["cloum4"] =s_JYbh;
                        dr["cloum5"] = s_Devid;
                        dr["cloum6"] =  ((double)在线时长 / 3600).ToString("0.0");//在线时长
                 
                        break;
                    default:
                        break;
                }
          



                dtreturns.Rows.Add(dr);
            }

    
            context.Response.Write(JSON.DatatableToDatatableJS(dtreturns, "www"));
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