using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    /// <summary>
    /// Dev_WorkState
    /// </summary>
    public partial class B_Dev_WorkState
    {
        private readonly DAL.D_Dev_WorkState dal = new DAL.D_Dev_WorkState();
        public B_Dev_WorkState()
        { }
        #region  BasicMethod
        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_Dev_WorkState model)
        {
            return dal.Add(model);
        }

        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_Dev_WorkState model)
        {
            return dal.Update(model);
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(int ID, string StateName)
        {

            return dal.Delete(ID, StateName);
        }

        /// <summary>
        /// 获得所有列表
        /// </summary>
        /// <returns></returns>
        public DataTable GetList()
        {

            return dal.GetList();
        }

                /// <summary>
        /// 获得所有列表  只要上下线
        /// </summary>
        /// <returns></returns>
        public DataTable GetListLog()
        {
            return dal.GetListLog();
        }



                /// <summary>
        /// 获得所有列表 维修统计
        /// </summary>
        /// <returns></returns>
        public DataTable GetListrEpairs()
        {
            return dal.GetListrEpairs();
        
        }

                /// <summary>
        /// 查找对应ID
        /// </summary>
        /// <returns></returns>
        public object GetListID(string StateName)
        {

            return dal.GetListID(StateName);
        }
     
        #endregion  BasicMethod

    }
}
