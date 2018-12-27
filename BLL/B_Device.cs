using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    /// <summary>
    /// Device
    /// </summary>
    public partial class Device
    {
        private readonly DAL.Device dal = new DAL.Device();
        public Device()
        { }
        #region  BasicMethod
        /// <summary>
        /// 是否存在该记录
        /// </summary>
        public bool Exists(string DevId)
        {
            return dal.Exists(DevId);
        }




               /// <summary>
        /// 是否存在该警员编号记录
        /// </summary>
        public bool ExistsJYBH(string JYBH)
        {

            return dal.ExistsJYBH(JYBH);
        
        }


                /// <summary>
        /// 是否存在该警员编号记录2
        /// </summary>
        public bool ExistsJYBH2(string JYBH, long ID)
        {

            return dal.ExistsJYBH2(JYBH, ID);
        }

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.Device model)
        {
            return dal.Add(model);
        }

        /// <summary>
        /// 更新指定数据数据
        /// </summary>
        public int UpdateData(Model.Device model)
        {

            return dal.UpdateData(model);
        }
        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.Device model)
        {
            return dal.Update(model);
        }


        public int UpdateRemind1(int DevType, string Manufacturer, string SBXH, string ProjName, DateTime CGSJ, DateTime BFQX, string Price, DateTime CreatDate, long ID, string TypeName, string Remind)
        {

            return dal.UpdateRemind1(DevType, Manufacturer, SBXH, ProjName, CGSJ, BFQX, Price, CreatDate, ID, TypeName, Remind);
        }


        
        /// <summary>
        /// 更新一条数据 设备提醒表
        /// </summary>
        public int UpdateRemind(Model.Device model)
        {
            return dal.UpdateRemind(model);
        }

        
        /// <summary>
        /// 更新WorkState设备状态 where 值为DevId设备编号
        /// </summary>
        public int UpdateWorkState_DevId(Model.Device model)
        {
            return dal.UpdateWorkState_DevId(model);
        }


            /// <summary>
        /// 更新WorkState设备状态
        /// </summary>
        public int UpdateWorkState(Model.Device model)
        {
            return dal.UpdateWorkState(model);
        
        }


          /// <summary>
        /// 更新AllocateState设备状态
        /// </summary>
        public int UpdateAllocateState(Model.Device model)
        {
            return dal.UpdateAllocateState(model);
        
        }


          /// <summary>
        /// 更新AllocateState设备状态 where 条件是DevId
        /// </summary>
        public int UpdateAllocateState_WDevId(string DevId)
        {
            return dal.UpdateAllocateState_WDevId(DevId);
        
        }



              /// <summary>
        /// 更新JYBH
        /// </summary>
        public int Update_JYBH(Model.Device model)
        {
            return dal.Update_JYBH(model);
        }

        /// <summary>
        /// 删除一条数据
        /// </summary>
        public bool Delete(long ID)
        {

            return dal.Delete(ID);
        }
       
        /// <summary>
        /// 得到一个对象实体
        /// </summary>
        public Model.Device GetModel(long ID)
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
        /// 获取设备总数量
        /// </summary>
        /// <returns></returns>
        public int getcount(int DeviceNmae, string strat, string now, int State, string three, string second, string search, int AllocateState)
        {
            return dal.getcount(DeviceNmae, strat, now, State, three, second, search, AllocateState);
        
        }


                /// <summary>
        /// 分页排序设备信息
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(int DeviceNmae, string strat, string now, int State, string three, string second, string search,int AllocateState, int startRowIndex, int maximumRows)
        {

            return dal.getcountpaging(DeviceNmae, strat, now, State, three, second, search, AllocateState,startRowIndex, maximumRows);
        }



        

        /// <summary>
        /// 导出设备登陆数据使用
        /// </summary>
        /// <param name="DeviceNmae"></param>
        /// <param name="strat"></param>
        /// <param name="now"></param>
        /// <param name="State"></param>
        /// <param name="three"></param>
        /// <param name="search"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="maximumRows"></param>
        /// <returns></returns>
        public DataTable OutExcel(int DeviceNmae, string strat, string now, int State, string three, string second, string search,int AllocateState)
        {

            return dal.OutExcel(DeviceNmae, strat, now, State, three, second, search, AllocateState);
        
        }


            /// <summary>
        /// 飘窗
        /// </summary>
        public DataTable BindBM(int DevType, string SJBM)
        {

            return dal.BindBM( DevType, SJBM);
        }


/// <summary>
        /// 导出设备提醒表
/// </summary>
/// <param name="Remind"></param>
/// <param name="DeviceNmae"></param>
/// <param name="strat"></param>
/// <param name="now"></param>
/// <param name="search"></param>
/// <returns></returns>
        public DataTable OutExcelRemind(int Remind, int DeviceNmae, string strat, string now, string three, string second, string search)
        {


            return dal.OutExcelRemind(Remind, DeviceNmae, strat, now, three,second, search);
        
        }


             /// <summary>
        /// 获取设备总数量 设备提醒
        /// </summary>
        /// <returns></returns>
        public int getcountRemind(int Remind, int DeviceNmae, string strat, string now, string three, string second, string search)
        {
            return dal.getcountRemind(Remind, DeviceNmae, strat, now, three, second, search);

        }


        
        /// <summary>
        /// 分页排序设备信息 设备提醒
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpagingRemind(int Remind, int DeviceNmae, string strat, string now, string three, string second, string search, int startRowIndex, int maximumRows)
        {

            return dal.getcountpagingRemind(Remind, DeviceNmae, strat, now, three, second, search, startRowIndex, maximumRows);
        }


      


        #endregion  BasicMethod
   
    }
}
