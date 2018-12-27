using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    /// <summary>
    /// OperationLog
    /// </summary>
    public partial class B_OperationLog
    {
        private readonly DAL.D_OperationLog dal = new DAL.D_OperationLog();
        public B_OperationLog()
        { }
        #region  BasicMethod
        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_OperationLog model)
        {
            return dal.Add(model);
        }

        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_OperationLog model)
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


        /// <summary>
        /// 得到一个对象实体
        /// </summary>
        public Model.M_OperationLog GetModel(long ID)
        {

            return dal.GetModel(ID);
        }

     

                /// <summary>
        /// 获取设备总数量 系统日志
        /// </summary>
        /// <returns></returns>
        public int getcount(string three, string second, string search)
        {
            return dal.getcount( three, second,search);

        }



           /// <summary>
        /// 分页排序设备信息 系统日志
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(string three, string second, string search, int startRowIndex, int maximumRows)
        {
            return dal.getcountpaging( three, second,search, startRowIndex, maximumRows);

        }

        #endregion  BasicMethod

    }
}
