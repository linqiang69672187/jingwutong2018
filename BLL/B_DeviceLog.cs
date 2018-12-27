using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    /// <summary>
    /// DeviceLog
    /// </summary>
    public partial class B_DeviceLog
    {
        private readonly DAL.D_DeviceLog dal = new DAL.D_DeviceLog();
        public B_DeviceLog()
        { }
        #region  BasicMethod
        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.M_DeviceLog model)
        {
            return dal.Add(model);
        }


        
        /// <summary>
        /// 增加一条数据 只增加BXR BZ LogTime 三格字段
        /// </summary>
        public int AddThree(Model.M_DeviceLog model)
        {


            return dal.AddThree(model);
        }


            /// <summary>
        /// 同意更新 只更新 BZ LogTime DevState BXR
        /// </summary>
        public int UpdateFour(Model.M_DeviceLog model)
        {
         return  dal.UpdateFour(model);
        
        }


        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_DeviceLog model)
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
        /// 得到一个对象实体
        /// </summary>
        public Model.M_DeviceLog GetModel(long ID)
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
        /// 导出维修统计
        /// </summary>
        /// <param name="three"></param>
        /// <param name="DeviceNmae"></param>
        /// <param name="strat"></param>
        /// <param name="now"></param>
        /// <param name="State"></param>
        /// <param name="search"></param>
        /// <returns></returns>
        public DataTable OutExcel(string three, string second, int DeviceNmae, string strat, string now, int State, string search)
        {

            return dal.OutExcel(three,second, DeviceNmae, strat, now,  State, search);
        }


        /// <summary>
        /// 获取设备总数量
        /// </summary>
        /// <returns></returns>
        public int getcount(string three, string second, int DeviceNmae, string strat, string now, int State, string search)
        {
            return dal.getcount(three,second, DeviceNmae, strat, now, State, search);

        }


        /// <summary>
        /// 分页排序设备信息
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(string three, string second, int DeviceNmae, string strat, string now, int State, string search, int startRowIndex, int maximumRows)
        {

            return dal.getcountpaging( three,second, DeviceNmae, strat, now, State,search, startRowIndex, maximumRows);
        }



               /// <summary>
        /// 导出回收统计
        /// </summary>
        /// <param name="three"></param>
        /// <param name="DeviceNmae"></param>
        /// <param name="strat"></param>
        /// <param name="now"></param>
        /// <param name="State"></param>
        /// <param name="search"></param>
        /// <returns></returns>

        public DataTable OutExcelRecycle(string three,string second, string strat, string now, int State, string search)
        {

            return dal.OutExcelRecycle( three,second, strat,now, State, search);

        }



        /// <summary>
        /// 获取设备总数量 回收表
        /// </summary>
        /// <returns></returns>
        public int getcountRecycle(string three, string second, string strat, string now, int State, string search)
        {

            return dal.getcountRecycle(three,second, strat, now, State, search);

        }



        /// <summary>
        /// 分页排序设备信息回收
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpagingRecycle(string three, string second, string strat, string now, int State, string search, int startRowIndex, int maximumRows)
        {
            return dal.getcountpagingRecycle(three, second,strat, now, State, search, startRowIndex, maximumRows);

        }



        
        /// <summary>
        /// 获取设备总数量 设备日志
        /// </summary>
        /// <returns></returns>
        public int getcountLog(int DeviceNmae, string three, string second, int State, string search)
        {
        
        return dal.getcountLog(DeviceNmae,three,second, State, search);
        
        }



        
        /// <summary>
        /// 分页排序设备信息 设备日志
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpagingLog(int DeviceNmae, string three, string second, int State, string search, int startRowIndex, int maximumRows)
        {
            return dal.getcountpagingLog(DeviceNmae, three, second,State, search, startRowIndex, maximumRows);

        }


        #endregion  BasicMethod
   
    }
}
