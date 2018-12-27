using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{/// <summary>
    /// Position
    /// </summary>
    public partial class Position
    {
        private readonly DAL.D_Position dal = new DAL.D_Position();
        public Position()
        { }
        #region  BasicMethod


                /// <summary>
        /// 获得所有列表
        /// </summary>
        /// <returns></returns>
        public DataTable GetList()
        {
            return dal.GetList();
        }

            /// <summary>
        /// 根据职务名称查找职务ID
        /// </summary>
        /// <returns></returns>
        public object GetListLDJB(string S_PositionName)
        {
            return dal.GetListLDJB(S_PositionName);
        }

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_Position model)
        {
            return dal.Add(model);
        }

        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_Position model)
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
        public Model.M_Position GetModel(int ID)
        {

            return dal.GetModel(ID);
        }

      
        #endregion  BasicMethod
 
    }
}
