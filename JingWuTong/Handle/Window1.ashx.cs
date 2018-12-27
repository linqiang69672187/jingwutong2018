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


namespace Policesystem.Handle
{
    /// <summary>
    /// getDataManagement 的摘要说明
    /// </summary>
    public class Window1 : IHttpHandler
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
            //string ssdd = context.Request.Form["ssdd"];
            string sszd = context.Request.Form["sszd"];
            string requesttype = context.Request.Form["requesttype"];
            //string search = context.Request.Form["search"];
            string sreachcondi = "";

            string URL = context.Request.Form["URL"];

            //int onlinevalue = int.Parse(context.Request.Form["onlinevalue"]) * 60;
            //int usedvalue = int.Parse(context.Request.Form["usedvalue"]) * 60;


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
            dtreturns.Columns.Add("cloum6",typeof(double));
            dtreturns.Columns.Add("cloum7");
            dtreturns.Columns.Add("cloum8", typeof(double));
            dtreturns.Columns.Add("cloum9");
            dtreturns.Columns.Add("cloum10");
            dtreturns.Columns.Add("cloum11");
            dtreturns.Columns.Add("cloum13", typeof(int));

            int days = Convert.ToInt16(context.Request.Form["dates"]);
            int statusvalue = 0;  //正常参考值
            int zxstatusvalue = 0;//在线参考值
            int devicescount = 0;  //汇总设备总数
            double zxsc = 0.0;  //汇总在线时长
            double spdx = 0.0;  //汇总视频大小
            Int64 cxl = 0;  //汇总查询量
            int wcxl = 0;   //无查询量设备数量
            int wcfl = 0;   //无处罚量设备数量
            int wsysb = 0;  //无使用设备数量
            int zxsb = 0; //在线设备



            int Allotment = 0;//配发设备数
            double Online = 0;//在线时长
            int use = 0;//设备使用数量
            double usagerate = 0;//设备使用率
            //辅警通和警务通新加
            int Auxiliary = 0;//辅警数
            double Illegally = 0;//违停采集
            double Punish = 0;//人均处罚量
            long Inquire = 0;//查询量
            double equipment = 0;//设备平均处罚量
            int Month=0;//本月无采集违停设备
             //执法记录仪
            int Unused = 0;//设备未使用数量
            double video = 0;//视频时长
            double Video_siz = 0;//视频大小


            string bmdm = ""; //汇总的部门代码
            int allstatu_device = 0;  //汇总使用率不为空数量
            string ddtitle;//大队标题

            var a = "";
            string isjujiguan;//判断是否是局机关
            string isxias="";
            //statusvalue = days * usedvalue;//超过10分钟算使用
            //zxstatusvalue = days * onlinevalue;//在线参考值

            allEntitys = SQLHelper.ExecuteRead(CommandType.Text, "SELECT BMDM,SJBM from [Entity] ", "11");

            DataTable Alarm_EveryDayInfo = null; //每日告警
            DataTable dUser = null;


                    bmdm = sszd;
                    ddtitle = context.Request.Form["sszdtext"];
                    switch (type)
                    {
                        case "5":
                            Alarm_EveryDayInfo = SQLHelper.ExecuteRead(CommandType.Text, "SELECT en.BMDM ,en.SJBM as [ParentID],us.XM as [Contacts],de.[DevId],ala.在线时长,ala.[AlarmType],ala.文件大小 from (SELECT [DevId],sum([VideLength]) as 在线时长,sum([FileSize]) as 文件大小,1 as AlarmType from [EveryDayInfo_ZFJLY]   where  [Time] >='" + begintime + "' and [Time] <='" + endtime + "'   group by [DevId] ) as ala left join [Device] as de on de.[DevId] = ala.[DevId] left join [Entity] as en on en.[BMDM] = de.[BMDM] left join ACL_USER as us on de.JYBH = us.JYBH  where  de.[DevType]=" + type + "" + isxias, "Alarm_EveryDayInfo");
                            break;
                        default:
                            Alarm_EveryDayInfo = SQLHelper.ExecuteRead(CommandType.Text, "SELECT en.BMDM ,en.SJBM as [ParentID],us.XM as [Contacts],de.[DevId],ala.在线时长,ala.[AlarmType],0 as 文件大小 from (SELECT [DevId],[AlarmType],sum([Value]) as 在线时长 from [Alarm_EveryDayInfo]   where [AlarmType] <> 6 and  [AlarmDay ] >='" + begintime + "' and [AlarmDay ] <='" + endtime + "'   group by [DevId],[AlarmType] ) as ala left join [Device] as de on de.[DevId] = ala.[DevId] left join [Entity] as en on en.[BMDM] = de.[BMDM] left join ACL_USER as us on de.JYBH = us.JYBH  where   de.[DevType]=" + type + "" + isxias, "Alarm_EveryDayInfo");
                            break;
                    }

                    if (sszd != "331000000000")
                    {

                        isxias = "   and ( en.SJBM='"+sszd+"')";
                    
                    }

                    if (type != "5")
                    {
                       
                        isjujiguan = "and A.BMDM<>'33100000000x'";
                    }
                    else
                    {
                        isjujiguan = "";
                    }

                    dtEntity = SQLHelper.ExecuteRead(CommandType.Text, " select * from"+
    "("+
   " select ntable.ID,ntable.Name,ntable.ParentID,ntable.Depth,ntable.Sort from  (SELECT A.BMDM as ID,A.BMQC as Name,A.SJBM as ParentID,A.BMJB AS Depth,A.Sort from [Entity] A  " +
    " where ( A.SJBM='" + sszd + "'"+isjujiguan+") and  A.BMQC IS NOT NULL AND A.BMQC <> '' ) as ntable " +

     " ) as nttable ORDER BY CASE WHEN Sort IS NULL THEN 1 ELSE Sort END desc", "2");



                    dUser = SQLHelper.ExecuteRead(CommandType.Text, "SELECT en.SJBM,us.BMDM FROM [dbo].[ACL_USER] us left join Entity en on us.BMDM = en.BMDM  ", "user");



            for (int i1 = 0; i1 < dtEntity.Rows.Count; i1++)
            {
       


                DataRow dr = dtreturns.NewRow();

                dr["cloum1"] = (i1 + 1).ToString();
                dr["cloum2"] = "<a href='"+URL+"?BMDM=" + dtEntity.Rows[i1]["ID"].ToString() + "&"+dtEntity.Rows[i1]["Name"].ToString()+"'>" + dtEntity.Rows[i1]["Name"].ToString() + "</a>";
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

                var entityids = GetSonID(dtEntity.Rows[i1]["ID"].ToString());
                List<string> strList = new List<string>();

                strList.Add(dtEntity.Rows[i1]["ID"].ToString());
    
                    foreach (entityStruct item in entityids)
                    {
                        strList.Add(item.BMDM);
                    }
                

                //var rows = from p in Alarm_EveryDayInfo.AsEnumerable()
                //           where (p.Field<string>("ParentID") == dtEntity.Rows[i1]["ID"].ToString()|| p.Field<string>("BMDM") == dtEntity.Rows[i1]["ID"].ToString())
                //           orderby p.Field<string>("DevId")
                //           select p;
        
              

                var rows = (from p in Alarm_EveryDayInfo.AsEnumerable()
                            where strList.ToArray().Contains(p.Field<string>("BMDM"))
                            orderby p.Field<string>("DevId")
                            select new dataStruct
                            {
                                BMDM = p.Field<string>("BMDM"),
                                ParentID = p.Field<string>("ParentID"),
                                在线时长 = p.Field<int>("在线时长"),
                                文件大小 = p.Field<int>("文件大小"),
                                AlarmType = p.Field<int>("AlarmType"),
                                DevId = p.Field<string>("DevId")
                            }).ToList<dataStruct>();


                //获得设备数量，及正常使用设备
                tmpRows = 0;
                foreach (dataStruct item in rows)
                {
          
                        switch (item.AlarmType.ToString())
                        {
                            case "1":
                                在线时长 += Convert.ToInt32(item.在线时长);
                                未使用 += ((Convert.ToInt32(item.在线时长) - statusvalue) <= 0) ? 1 : 0;
                                在线 += ((Convert.ToInt32(item.在线时长) - zxstatusvalue) > 0) ? 1 : 0;
                                文件大小 += Convert.ToInt32(item.文件大小);
                                break;
                            case "2":
                                处理量 += Convert.ToInt32(item.在线时长);
                                无处罚量 += (Convert.ToInt32(item.在线时长) == 0) ? 1 : 0;
                                break;
                            case "5":
                                查询量 += Convert.ToInt32(item.在线时长);
                                无查询量 += (Convert.ToInt32(item.在线时长) == 0) ? 1 : 0;
                                break;
                        }
                        if (item.DevId.ToString() != tmpDevid)
                        {
                            tmpRows += 1;  //新设备ID不重复
                            tmpDevid = item.DevId.ToString();
                            status += (Convert.ToInt32(item.在线时长) - statusvalue > 0) ? 1 : 0;
                            int c = (Convert.ToInt32(item.在线时长) - statusvalue > 0) ? 1 : 0;
                            if ( c==0)
                            {
                                a = "3435";
                            }
                            allstatu_device += (Convert.ToInt32(item.在线时长) - statusvalue > 0) ? 1 : 0;
                        }
                     
                }

                tmpList.Clear();

                var userrows = from p in dUser.AsEnumerable()
                           where (p.Field<string>("SJBM") == dtEntity.Rows[i1]["ID"].ToString()|| p.Field<string>("BMDM") == dtEntity.Rows[i1]["ID"].ToString()) select p;
                usercount = userrows.Count();

                if (usercount == 0)//规避usercount为0后面计算错误
                {

                    usercount = 1;
                }

                int countdevices = tmpRows;
                double deviceuse = Math.Round((double)status * 100 / (double)countdevices,2);

                dr["cloum3"] = countdevices;//配发台数
                devicescount += countdevices;
                Allotment += countdevices;//汇总配发台数

                switch (type)
                {
                    case "4":
                    case "6":
                        dr["cloum4"] = usercount; //辅警数
                        Auxiliary += usercount;//辅警数汇总
                        zxsc += 处理量;
                        dr["cloum5"] = 处理量;//违停采集
                        Illegally += double.Parse(处理量.ToString());//违停采集汇总

                        dr["cloum6"] = Math.Round((double)处理量 / usercount, 2);//人均处罚量
                        //Punish+=(double)Math.Round((double)处理量 / usercount, 2);//人均处罚量汇总

                        dr["cloum7"] = 查询量;
                        Inquire += 查询量;//查询量汇总

                        dr["cloum8"] =  (countdevices == 0) ? 0 : Math.Round((double)处理量 / countdevices, 2);;//设备处罚量

                       // equipment+=(double) Math.Round((double)countdevices / usercount, 2);//设备处罚量汇总

                        dr["cloum10"] = 无处罚量;
                        Month += 无处罚量;//本月无采集违停设备汇总

                        break;
                    case "5":   //执法记录仪
                     dr["cloum4"] =status;//设备使用数量
                        use +=status;//汇总设备使用数量
                         dr["cloum5"] = 未使用;//设备未使用数量
                         Unused += 未使用;//设备未使用数量汇总
                         dr["cloum6"] = ((double)在线时长 / 3600).ToString("0.00");//视频时长总和
                         video +=double.Parse(((double)在线时长 / 3600).ToString("0.00"));//视频时长总和汇总
                         dr["cloum7"] = ((double)文件大小 / 1048576).ToString("0.00");;//视频大小（GB）
                         Video_siz += double.Parse( ((double)文件大小 / 1048576).ToString("0.00"));//视频大小（GB)汇总
                         dr["cloum8"] = (countdevices != 0) ? (deviceuse) : 0;//设备使用率
                         zxsc += (double)在线时长 / 3600;
                          //usagerate +=(double)((countdevices != 0) ? (deviceuse) : 0);//汇总设备使用率
                        break;
                    case "1":
                    case "2":
                    case "3":
                    case "7":
                        dr["cloum4"] =status;//设备使用数量
                        use +=status;//汇总设备使用数量
                        dr["cloum5"] = ((double)在线时长 / 3600).ToString("0.00");;//在线时长
                        Online += double.Parse(((double)在线时长 / 3600).ToString("0.00"));//汇总在线时长
                      
                        dr["cloum6"] = (countdevices != 0) ? (deviceuse):0.0;//设备使用率
                        //usagerate +=(double)((countdevices != 0) ? (deviceuse) : 0.0);//汇总设备使用率
                        zxsc += (double)在线时长 / 3600;

                        break;
                    default:
                        break;
                }
                //zxsb += 在线;
                //dr["cloum14"] = 在线;
                //dr["cloum12"] = dtEntity.Rows[i1]["ID"].ToString();
               // dr["cloum6"] = (countdevices != 0) ? (deviceuse):0;
   


                dtreturns.Rows.Add(dr);
            }



            var Fcloum = "";
            var Scloum = "";

            switch (type)
            {
                case "4":
                case "6":

                    Fcloum = "cloum8";
                    Scloum = "cloum9";
                    break;
                case "5":   //执法记录仪

                    Fcloum = "cloum8";
                    Scloum = "cloum9";
                    break;
                case "1":
                case "2":
                case "3":
                case "7":
                    Fcloum = "cloum6";
                    Scloum = "cloum7";

                    break;
                default:
                    break;

            }

            int orderno = 1;
            var query = (from p in dtreturns.AsEnumerable()
                         orderby p.Field<double>(Fcloum) descending
                         select p) as IEnumerable<DataRow>;

            double temsyl = 0.0;
            int temorder = 1;
            foreach (var item in query)
            {
                if (temsyl == double.Parse(item[Fcloum].ToString()))
                {
                    item[Scloum] = temorder;

                }
                else
                {
                    item[Scloum] = orderno;
                    temsyl = double.Parse((item[Fcloum].ToString()));
                    temorder = orderno;
                }
                orderno += 1;
            }
            query = query.OrderBy(p => p["cloum13"]);
            dtreturns = query.CopyToDataTable<DataRow>();

            DataRow dr2 = dtreturns.NewRow();
            dr2["cloum1"] = dtreturns.Rows.Count + 1;
            dr2["cloum2"] = "合计";//ddtitle;
            dr2["cloum3"] = devicescount;

            switch (type)
            {
                case "4":
                case "6":
                    dr2["cloum3"] = Allotment;
                    dr2["cloum4"] = Auxiliary; //辅警数
                    dr2["cloum5"] = Illegally;//违停采集
                    dr2["cloum6"] = (Auxiliary == 0) ? "0" : (zxsc / Auxiliary).ToString("0.00");//人均处罚量
                    dr2["cloum7"] = Inquire;
                    dr2["cloum8"] = (devicescount == 0) ? "0" : (zxsc / devicescount).ToString("0.00");//设备处罚量
                    dr2["cloum9"] = "/";
                    dr2["cloum10"] = Month;

                    break;
                case "5":   //执法记录仪
                    dr2["cloum3"] = Allotment;
                    dr2["cloum4"] = use;//设备使用数量
                    dr2["cloum5"] = Unused;//设备未使用数量
                    dr2["cloum6"] = video;//视频时长总和
                    dr2["cloum7"] = Video_siz;//视频大小（GB）
                    usagerate = (devicescount == 0) ? 0 : ((double)allstatu_device * 100 / devicescount);
                    dr2["cloum8"] = Math.Round(usagerate, 2);//设备使用率
                    dr2["cloum9"] = "/";

                    break;
                case "1":
                case "2":
                case "3":
                case "7":
                    dr2["cloum3"] = Allotment;
                    dr2["cloum4"] = use;
                    dr2["cloum5"] = Online;
                    usagerate = (devicescount == 0) ? 0 : ((double)allstatu_device * 100 / devicescount);
                    dr2["cloum6"] = Math.Round(usagerate, 2);
                    dr2["cloum7"] = "/";

                    break;
                default:
                    break;
            }
            dtreturns.Rows.Add(dr2);



            //}
            context.Response.Write(JSON.DatatableToDatatableJS(dtreturns,"www"));
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