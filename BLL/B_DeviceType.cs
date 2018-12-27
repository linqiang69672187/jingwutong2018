using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    /// <summary>
    /// DeviceType
    /// </summary>
    public partial class B_DeviceType
    {
        private readonly DAL.D_DeviceType dal = new DAL.D_DeviceType();
        public B_DeviceType()
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
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_DeviceType model)
        {
            return dal.Add(model);
        }

        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_DeviceType model)
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
        public Model.M_DeviceType GetModel(int ID)
        {

            return dal.GetModel(ID);
        }


        /// <summary>
        /// 模糊查询
        /// </summary>
        /// <returns></returns>
        public DataTable GetLike(string field, string fieldvalue)
        {

            return dal.GetLike(field, fieldvalue);
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
        /// 查找对应ID
        /// </summary>
        /// <returns></returns>
        public object GetListID(string TypeName)
        {

            return dal.GetListID(TypeName);
        }

        #endregion  BasicMethod
 
    }
}
