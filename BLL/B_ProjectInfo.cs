using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    /// <summary>
    /// ProjectInfo
    /// </summary>
    public partial class B_ProjectInfo
    {
        private readonly DAL.D_ProjectInfo dal = new DAL.D_ProjectInfo();
        public B_ProjectInfo()
        { }
        #region  BasicMethod
        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_ProjectInfo model)
        {
            return dal.Add(model);
        }

        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int  Update(Model.M_ProjectInfo model)
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
        public Model.M_ProjectInfo GetModel(string project)
        {

            return dal.GetModel(project);
        }



        /// <summary>
        /// 获得所有列表
        /// </summary>
        /// <returns></returns>
        public DataTable GetList()
        {

            return dal.GetList();
        }

       
        #endregion  BasicMethod
       
    }
}
