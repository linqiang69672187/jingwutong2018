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
            DataTable pages = SQLHelper.ExecuteRead(CommandType.Text, "SELECT [id],[name] ,[JB],[Sort],[parent_id] FROM [Pages] ", "pages");
            DataTable buttons = SQLHelper.ExecuteRead(CommandType.Text, "SELECT [id],[name],[Sort],[page_name] FROM [Buttons]", "buttons");
            OrderedEnumerableRowCollection<DataRow> pagesrows;
            pagesrows = from p in pages.AsEnumerable()
                   where (p.Field<int>("JB") == 0)
                   orderby p.Field<int>("Sort") ascending
                   select p;
            StringBuilder retJson = new StringBuilder();


       
            retJson.Append("[");
            int h = 0;
            foreach (var entityitem in pagesrows)
                {
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