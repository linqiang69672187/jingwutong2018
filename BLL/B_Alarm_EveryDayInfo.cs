using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    /// <summary>
    /// Alarm_EveryDayInfo
    /// </summary>
    public partial class B_Alarm_EveryDayInfo
    {
        private readonly DAL.D_Alarm_EveryDayInfo dal = new DAL.D_Alarm_EveryDayInfo();
        public B_Alarm_EveryDayInfo()
        { }
        #region  BasicMethod

        
        /// <summary>
        /// 飘窗
        /// </summary>
        public DataTable Bind_Alarm_EveryDayInfo(int DevType, string SJBM, DateTime Start_Time, DateTime End_Time)
        {
            return dal.Bind_Alarm_EveryDayInfo(DevType, SJBM, Start_Time, End_Time);


        }


        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_Alarm_EveryDayInfo model)
        {
            return dal.Add(model);
        }

        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_Alarm_EveryDayInfo model)
        {
            return dal.Update(model);
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(long ID)
        {

            return dal.Delete(ID);
        }

        /// <summary>
        /// 得到一个对象实体
        /// </summary>
        public Model.M_Alarm_EveryDayInfo GetModel(long ID)
        {

            return dal.GetModel(ID);
        }

        

        #endregion  BasicMethod

    }
}
