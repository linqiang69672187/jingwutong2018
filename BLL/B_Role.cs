using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    /// <summary>
    /// Role
    /// </summary>
    public partial class B_Role
    {
        private readonly DAL.D_Role dal = new DAL.D_Role();
        public B_Role()
        { }
        #region  BasicMethod
        /// <summary>
        /// 是否存在该记录
        /// </summary>
        public object Exists(int ID)
        {
            return dal.Exists(ID);
        }


        
        /// <summary>
        /// 根据角色名称查找角色ID
        /// </summary>
        /// <param name="JSIDName"></param>
        /// <returns></returns>
        public object ExistsJSID(string JSIDName)
        {
            return dal.ExistsJSID(JSIDName);

        }

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_Role model)
        {
            return dal.Add(model);
        }

        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_Role model)
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
        public Model.M_Role GetModel(int ID)
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


        
        /// <summary>
        /// 获取设备总数量
        /// </summary>
        /// <returns></returns>
        public int getcount(string search)
        {
            return dal.getcount(search);
        
        }


                /// <summary>
        /// 分页排序设备信息
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(string search, int startRowIndex, int maximumRows)
        {
            return dal.getcountpaging(search, startRowIndex, maximumRows);
        
        }


        #endregion  BasicMethod

    }
}
