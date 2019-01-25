using DbComponent;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;

namespace JingWuTong.Handle
{
    /// <summary>
    /// permissions_set 的摘要说明
    /// </summary>
    public class permissions_set : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "text/plain";
            string type = context.Request.Form["requesttype"];
            DataTable rolo_power=null;
            DataTable rolo = null;
            if (context.Request.Form["roleid"]!="")
            {
                rolo_power = SQLHelper.ExecuteRead(CommandType.Text, "SELECT page_or_buttons_id,type,enable FROM [role_power]   where role_id =" + context.Request.Form["roleid"], "rolo_power");
                rolo = SQLHelper.ExecuteRead(CommandType.Text, "SELECT id,RoleName,Bz,Crateator FROM [role]   where id =" + context.Request.Form["roleid"], "rolo");
            }

            StringBuilder retJson = new StringBuilder();

            switch (type)
            {
                case "add":
                    goto add;
                case "save":
                    goto save;
   
            }
            add:
            DataTable pages = SQLHelper.ExecuteRead(CommandType.Text, "SELECT id,name,JB,Sort,parent_id FROM Pages ", "pages");
            DataTable buttons = SQLHelper.ExecuteRead(CommandType.Text, "SELECT id,name,Sort,page_name FROM Buttons", "buttons");
            OrderedEnumerableRowCollection<DataRow> pagesrows;
            OrderedEnumerableRowCollection<DataRow> buttonsrows;
            OrderedEnumerableRowCollection<DataRow> childpagesrows;
            pagesrows = from p in pages.AsEnumerable()
                   where (p.Field<int>("JB") == 0)
                   orderby p.Field<int>("Sort") ascending
                   select p;
            retJson.Append("{\"");
            retJson.Append("pages");
            retJson.Append('"');
            retJson.Append(":[");
            int h = 0;
            int n = 0;
            int i = 0;
            foreach (var fPage in pagesrows)
                {
                buttonsrows = from p in buttons.AsEnumerable()
                              where (p.Field<string>("page_name") == fPage["name"].ToString())
                              orderby p.Field<int>("Sort") ascending
                              select p;
                childpagesrows = from p in pages.AsEnumerable()
                              where (p.Field<int>("parent_id") == int.Parse(fPage["id"].ToString()))
                              orderby p.Field<int>("Sort") ascending
                              select p;

                if (h != 0) retJson.Append(',');
                    h++;
                    retJson.Append('{');
                    retJson.Append('"');
                    retJson.Append("name");
                    retJson.Append('"');
                    retJson.Append(":");
                    retJson.Append('"');
                    retJson.Append(fPage["name"].ToString());
                    retJson.Append('"');
                    retJson.Append(',');
                    retJson.Append('"');
                    retJson.Append("id");
                    retJson.Append('"');
                    retJson.Append(":");
                    retJson.Append('"');
                    retJson.Append(fPage["id"].ToString());
                    retJson.Append('"');
                    retJson.Append(',');
                    retJson.Append('"');
                  
                    retJson.Append("ischecked");
                    retJson.Append('"');
                    retJson.Append(":");
                    if (rolo_power != null)
                    {
                     
                   var powerrows = from p in rolo_power.AsEnumerable()
                                  where (p.Field<int>("page_or_buttons_id") == int.Parse(fPage["id"].ToString()) && p.Field<string>("type") == "page" && p.Field<Boolean>("enable")==true)
                                  select p;
                    retJson.Append((powerrows.Count()>0)?"true":"false");
                    }
                    else
                    {
                        retJson.Append("false");
                    }
                
                    retJson.Append(',');
                    retJson.Append('"');
                    retJson.Append("buttons");
                    retJson.Append('"');
                    retJson.Append(":");
                    retJson.Append("[");
                     n = 0;

                        foreach (var btitem in buttonsrows)
                        {
                            if (n != 0) retJson.Append(',');
                            n++;
                            retJson.Append('{');
                            retJson.Append('"');
                            retJson.Append("name");
                            retJson.Append('"');
                            retJson.Append(":");
                            retJson.Append('"');
                            retJson.Append(btitem["name"].ToString());
                            retJson.Append('"');
                            retJson.Append(',');
                            retJson.Append('"');
                            retJson.Append("id");
                            retJson.Append('"');
                            retJson.Append(":");
                            retJson.Append('"');
                            retJson.Append(btitem["id"].ToString());
                            retJson.Append('"');
                            retJson.Append(',');
                            retJson.Append('"');
                            retJson.Append("ischecked");
                            retJson.Append('"');
                            retJson.Append(":");
                        if (rolo_power != null)
                        {

                            var powerrows = from p in rolo_power.AsEnumerable()
                                            where (p.Field<int>("page_or_buttons_id") == int.Parse(btitem["id"].ToString()) && p.Field<string>("type") == "button" && p.Field<Boolean>("enable") == true)
                                            select p;
                            retJson.Append((powerrows.Count() > 0) ? "true" : "false");
                        }
                        else
                        {
                            retJson.Append("false");
                        }
                    retJson.Append('}');
                         }
                    retJson.Append("]");
                    retJson.Append(',');
                    retJson.Append('"');
                    retJson.Append("child_page");
                    retJson.Append('"');
                    retJson.Append(":");
                    retJson.Append("[");
                      i = 0;
                     foreach (var childitem in childpagesrows)
                        {
                              buttonsrows = from p in buttons.AsEnumerable()
                                  where (p.Field<string>("page_name") == childitem["name"].ToString())
                                  orderby p.Field<int>("Sort") ascending
                                  select p;
                    if (i != 0) retJson.Append(',');
                            i++;
                            retJson.Append('{');
                            retJson.Append('"');
                            retJson.Append("name");
                            retJson.Append('"');
                            retJson.Append(":");
                            retJson.Append('"');
                            retJson.Append(childitem["name"].ToString());
                            retJson.Append('"');
                            retJson.Append(',');
                            retJson.Append('"');
                            retJson.Append("id");
                            retJson.Append('"');
                            retJson.Append(":");
                            retJson.Append('"');
                            retJson.Append(childitem["id"].ToString());
                            retJson.Append('"');
                            retJson.Append(',');
                            retJson.Append('"');
                            retJson.Append("ischecked");
                            retJson.Append('"');
                            retJson.Append(":");

                            if (rolo_power != null)
                            {

                                var powerrows = from p in rolo_power.AsEnumerable()
                                                where (p.Field<int>("page_or_buttons_id") == int.Parse(childitem["id"].ToString()) && p.Field<string>("type") == "page" && p.Field<Boolean>("enable") == true)
                                                select p;
                                retJson.Append((powerrows.Count() > 0) ? "true" : "false");
                            }
                            else
                            {
                                retJson.Append("false");
                            }
                            retJson.Append(',');
                            retJson.Append('"');
                            retJson.Append("buttons");
                            retJson.Append('"');
                            retJson.Append(":");
                            retJson.Append("[");
                            n = 0;
                                    foreach (var btitem in buttonsrows)
                                    {
                                        if (n != 0) retJson.Append(',');
                                        n++;
                                        retJson.Append('{');
                                        retJson.Append('"');
                                        retJson.Append("name");
                                        retJson.Append('"');
                                        retJson.Append(":");
                                        retJson.Append('"');
                                        retJson.Append(btitem["name"].ToString());
                                        retJson.Append('"');
                                        retJson.Append(',');
                                        retJson.Append('"');
                                        retJson.Append("id");
                                        retJson.Append('"');
                                        retJson.Append(":");
                                        retJson.Append('"');
                                        retJson.Append(btitem["id"].ToString());
                                        retJson.Append('"');
                                        retJson.Append(',');
                                        retJson.Append('"');
                                        retJson.Append("ischecked");
                                        retJson.Append('"');
                                        retJson.Append(":");

                                        if (rolo_power != null)
                                        {

                                            var powerrows = from p in rolo_power.AsEnumerable()
                                                            where (p.Field<int>("page_or_buttons_id") == int.Parse(btitem["id"].ToString()) && p.Field<string>("type") == "button" && p.Field<Boolean>("enable") == true)
                                                            select p;
                                            retJson.Append((powerrows.Count() > 0) ? "true" : "false");
                                        }
                                        else
                                        {
                                            retJson.Append("false");
                                        }
                                        retJson.Append('}');
                                    }

                           retJson.Append("]");
                           retJson.Append('}');
                        }
                retJson.Append("]");
                retJson.Append('}');

            }
            retJson.Append("],\"role\":{");
            retJson.Append('"');
            retJson.Append("name");
            retJson.Append('"');
            retJson.Append(":");
            retJson.Append('"');
            if (rolo != null)
            {
                retJson.Append(rolo.Rows[0]["RoleName"].ToString());
            }
            retJson.Append('"');
            retJson.Append(',');
            retJson.Append('"');
            retJson.Append("id");
            retJson.Append('"');
            retJson.Append(":");
            retJson.Append('"');
            if (rolo != null)
            {
                retJson.Append(rolo.Rows[0]["id"].ToString());
            }
            retJson.Append('"');
            retJson.Append(',');
            retJson.Append('"');
            retJson.Append("remark");
            retJson.Append('"');
            retJson.Append(":");
            retJson.Append('"');
            if (rolo != null)
            {
                retJson.Append(rolo.Rows[0]["Bz"].ToString());
            }
            retJson.Append('"');
            retJson.Append(',');
            retJson.Append('"');
            retJson.Append("creater");
            retJson.Append('"');
            retJson.Append(":");
            retJson.Append('"');
            if (rolo!=null)
            {
                retJson.Append(rolo.Rows[0]["crateator"].ToString());
            }
            retJson.Append('"');


            retJson.Append("}}");

            context.Response.Write(retJson.ToString());
            return;


            save://添加标签
            string data = context.Request.Form["data"];
            string role = context.Request.Form["role"];
            List<page> pagesdata = ReadingJson(data);
            role roledata = ReadingRoleJson(role);

            SqlParameter[] sp1 = new SqlParameter[4];
            sp1[0] = new SqlParameter("@id", roledata.id);
            sp1[1] = new SqlParameter("@Name", roledata.Name);
            sp1[2] = new SqlParameter("@creater", roledata.creater);
            sp1[3] = new SqlParameter("@remark", roledata.remark);
            var id= SQLHelper.ExecuteScalar(CommandType.Text, "begin tran update Role with (serializable) set RoleName=@Name,Crateator=@creater,Bz=@remark where id = @id  if @@rowcount = 0 begin insert into Role (RoleName,Crateator,Bz) values (@Name,@creater,@remark) select SCOPE_IDENTITY() end commit tran", sp1);

            if (id == null) id = roledata.id;
            try { 
            foreach (var pgitem in pagesdata)
            {
                SqlParameter[] pagesp = new SqlParameter[4];
                pagesp[0] = new SqlParameter("@page_or_buttons_id", pgitem.id);
                pagesp[1] = new SqlParameter("@enable", pgitem.ischecked);
                pagesp[2] = new SqlParameter("@role_id", id);
                pagesp[3] = new SqlParameter("@type", "page");
                SQLHelper.ExecuteNonQuery(CommandType.Text, "begin tran update role_power with (serializable) set enable=@enable where page_or_buttons_id = @page_or_buttons_id and role_id=@role_id and  @type=type if @@rowcount = 0 begin insert into role_power (page_or_buttons_id,enable,role_id,type) values (@page_or_buttons_id,@enable,@role_id,@type) end commit tran", pagesp);
                foreach (var bt in pgitem.buttons)
                {
                    SqlParameter[] btsp = new SqlParameter[4];
                    btsp[0] = new SqlParameter("@page_or_buttons_id", bt.id);
                    btsp[1] = new SqlParameter("@enable", bt.ischecked);
                    btsp[2] = new SqlParameter("@role_id", id);
                    btsp[3] = new SqlParameter("@type", "button");
                    SQLHelper.ExecuteNonQuery(CommandType.Text, "begin tran update role_power with (serializable) set enable=@enable where  page_or_buttons_id = @page_or_buttons_id and role_id=@role_id and  @type=type  if @@rowcount = 0 begin insert into role_power (page_or_buttons_id,enable,role_id,type) values (@page_or_buttons_id,@enable,@role_id,@type) end commit tran", btsp);
                }

                foreach (var pg in pgitem.child_page)
                {
                    SqlParameter[] child_pagesp = new SqlParameter[4];
                    child_pagesp[0] = new SqlParameter("@page_or_buttons_id", pg.id);
                    child_pagesp[1] = new SqlParameter("@enable", pg.ischecked);
                    child_pagesp[2] = new SqlParameter("@role_id", id);
                    child_pagesp[3] = new SqlParameter("@type", "page");
                    SQLHelper.ExecuteNonQuery(CommandType.Text, "begin tran update role_power with (serializable) set enable=@enable where  page_or_buttons_id = @page_or_buttons_id and role_id=@role_id and  @type=type  if @@rowcount = 0 begin insert into role_power (page_or_buttons_id,enable,role_id,type) values (@page_or_buttons_id,@enable,@role_id,@type) end commit tran", child_pagesp);
                    foreach (var bt in pg.buttons)
                    {
                        SqlParameter[] btsp = new SqlParameter[4];
                        btsp[0] = new SqlParameter("@page_or_buttons_id", bt.id);
                        btsp[1] = new SqlParameter("@enable", bt.ischecked);
                        btsp[2] = new SqlParameter("@role_id", id);
                        btsp[3] = new SqlParameter("@type", "button");
                        SQLHelper.ExecuteNonQuery(CommandType.Text, "begin tran update role_power with (serializable) set enable=@enable where  page_or_buttons_id = @page_or_buttons_id and role_id=@role_id and  @type=type  if @@rowcount = 0 begin insert into role_power (page_or_buttons_id,enable,role_id,type) values (@page_or_buttons_id,@enable,@role_id,@type) end commit tran", btsp);
                    }




                }



            }
            }
            catch (Exception e) {

            }
            context.Response.Write("1");
            return;


        }
        /// <summary>
        ///     读取json
        /// </summary>
        /// <param name="desktopPath"></param>
        private static List<page> ReadingJson(string str)
        {

            //转换
            var jArray = JsonConvert.DeserializeObject<List<page>>(str);
            return jArray;
        }
        private static role ReadingRoleJson(string str)
        {

            //转换
            var jArray = JsonConvert.DeserializeObject<role>(str);
            return jArray;
        }
        public class page
        {
            public int id { get; set; }
            public string Name { get; set; }
            public Boolean ischecked { get; set; }
            public List<page> child_page { get; set; }
            public List<button> buttons { get; set; }

        }
        public class role
        {
            public int id { get; set; }
            public string Name { get; set; }
            public string remark { get; set; }
            public string creater { get; set; }


        }
        public class button
        {
            public int id { get; set; }
            public string Name { get; set; }
            public Boolean ischecked { get; set; }
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