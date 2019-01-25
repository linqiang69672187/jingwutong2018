using DbComponent;
using System;
using System.Collections.Generic;
using System.Data;
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
            string type = "新增";
            DataTable pages = SQLHelper.ExecuteRead(CommandType.Text, "SELECT id,name,JB,Sort,parent_id FROM Pages ", "pages");
            DataTable buttons = SQLHelper.ExecuteRead(CommandType.Text, "SELECT id,name,Sort,page_name FROM Buttons", "buttons");
            OrderedEnumerableRowCollection<DataRow> pagesrows;
            OrderedEnumerableRowCollection<DataRow> buttonsrows;
            OrderedEnumerableRowCollection<DataRow> childpagesrows;
            pagesrows = from p in pages.AsEnumerable()
                   where (p.Field<int>("JB") == 0)
                   orderby p.Field<int>("Sort") ascending
                   select p;
            StringBuilder retJson = new StringBuilder();


       
            retJson.Append("[");
            int h = 0;
            int n = 0;
            int i = 0;
            foreach (var entityitem in pagesrows)
                {
                buttonsrows = from p in buttons.AsEnumerable()
                              where (p.Field<string>("page_name") == entityitem["name"].ToString())
                              orderby p.Field<int>("Sort") ascending
                              select p;
                childpagesrows = from p in pages.AsEnumerable()
                              where (p.Field<int>("parent_id") == int.Parse(entityitem["id"].ToString()))
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
                    retJson.Append(entityitem["name"].ToString());
                    retJson.Append('"');
                    retJson.Append(',');
                    retJson.Append('"');
                    retJson.Append("checked");
                    retJson.Append('"');
                    retJson.Append(":");
                    retJson.Append("false");
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
                            retJson.Append("checked");
                            retJson.Append('"');
                            retJson.Append(":");
                            retJson.Append("false");
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
                            retJson.Append("checked");
                            retJson.Append('"');
                            retJson.Append(":");
                            retJson.Append("false");
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
                                        retJson.Append("checked");
                                        retJson.Append('"');
                                        retJson.Append(":");
                                        retJson.Append("false");
                                        retJson.Append('}');
                                    }

                           retJson.Append("]");
                           retJson.Append('}');
                        }
                retJson.Append("]");
                retJson.Append('}');

            }
            retJson.Append("]");
            context.Response.Write(retJson.ToString());

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