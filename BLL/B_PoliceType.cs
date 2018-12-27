using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    /// <summary>
    /// PoliceType
    /// </summary>
    public partial class B_PoliceType
    {
        private readonly DAL.D_PoliceType dal = new DAL.D_PoliceType();
        public B_PoliceType()
        { }
        #region  BasicMethod
        /// <summary>
        /// 是否存在该记录
        /// </summary>
        public bool Exists(int ID)
        {
            return dal.Exists(ID);
        }


                /// <summary>
        /// 根据警员名称查找警员相对应类型
        /// </summary>
        /// <param name="JYLXName"></param>
        /// <returns></returns>
        public object ExistsJYLX(string JYLXName)
        {
            return dal.ExistsJYLX(JYLXName);

        }

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_PoliceType model)
        {
            return dal.Add(model);
        }

        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_PoliceType model)
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
        public Model.M_PoliceType GetModel(int ID, string TypeName)
        {

            return dal.GetModel(ID);
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
