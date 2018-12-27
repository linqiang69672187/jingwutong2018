using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    /// <summary>
    /// AlarmGate
    /// </summary>
    public partial class B_AlarmGate
    {
        private readonly DAL.D_AlarmGate dal = new DAL.D_AlarmGate();
        public B_AlarmGate()
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
        public int Add(Model.M_AlarmGate model)
        {
            return dal.Add(model);
        }

        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(int CommonAlarmGate, int UrgencyAlarmGate, string TypeName, string Name, int ID)
        {
            return dal.Update( CommonAlarmGate, UrgencyAlarmGate, TypeName, Name, ID);
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
        public Model.M_AlarmGate GetModel(int ID)
        {

            return dal.GetModel(ID);
        }


        
        /// <summary>
        /// 获取告警总数
        /// </summary>
        /// <returns></returns>
        public int getcount(int DeviceNmae)
        {
            return dal.getcount(DeviceNmae);
        }

       /// <summary>
        /// 分页排序告警
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(int DeviceNmae, int startRowIndex, int maximumRows)
        {
            return dal.getcountpaging(DeviceNmae, startRowIndex, maximumRows);

        }


        #endregion  BasicMethod
   
    }
}
