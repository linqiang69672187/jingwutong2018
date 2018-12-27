using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    /// <summary>
    /// ProjectDetail
    /// </summary>
    public partial class B_ProjectDetail
    {
        private readonly DAL.D_ProjectDetail dal = new DAL.D_ProjectDetail();
        public B_ProjectDetail()
        { }
        #region  BasicMethod
        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_ProjectDetail model)
        {
            return dal.Add(model);
        }

        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_ProjectDetail model)
        {
            return dal.Update(model);
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(int ID)
        {

            return dal.Delete(ID);
        }


        /// <summary>
        /// 得到一个对象实体
        /// </summary>
        public Model.M_ProjectDetail GetModel(int ID)
        {

            return dal.GetModel(ID);
        }


        /// <summary>
        /// 获得所有列表
        /// </summary>
        /// <returns></returns>
        public DataTable GetList(string XMBH)
        {

            return dal.GetList(XMBH);
        }


        #endregion  BasicMethod

    }
}
