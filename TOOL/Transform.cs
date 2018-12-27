using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TOOL
{
   public static class Transform
    {
       /// <summary>
        /// 此方法判断要传递的参数是否为null， 如果为Null, 则返回值DBNLL.Value,主要用户网数据库添加或者更新数据
       /// </summary>
       /// <param name="obj"></param>
       /// <returns></returns>
       public static object CheckIsNull(object obj)  
       {
           if (obj == null)
           {
               return DBNull.Value;
           }
           else
           {
               return obj;
           }
       }

       /// <summary>
       /// 此方法是从数据库中读取数据，如果数据库中的数据为DBNull.Value, 则返回null
       /// </summary>
       /// <param name="obj"></param>
       /// <returns></returns>
       public static object CheckIsDbNull(object obj)    
       {
           if (Convert.IsDBNull(obj))
           {
               return null;
           }
           else
           {
               return obj;
           }
       }


    }
}
