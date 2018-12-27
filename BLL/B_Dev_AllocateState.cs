using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{/// <summary>
    /// Dev_AllocateState
    /// </summary>
    public partial class B_Dev_AllocateState
    {
        private readonly DAL.D_Dev_AllocateState dal = new DAL.D_Dev_AllocateState();
        public B_Dev_AllocateState()
        { }
        #region  BasicMethod
        /// <summary>
        ///加载下拉列表
        /// </summary>
        public DataTable GetList()
        {
            return dal.GetList();
        }

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_Dev_AllocateState model)
        {
            return dal.Add(model);
        }

        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_Dev_AllocateState model)
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

      
        #endregion  BasicMethod

    }
}
