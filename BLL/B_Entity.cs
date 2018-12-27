using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{

    /// <summary>
    /// Entity
    /// </summary>
    public partial class Entity
    {
        private readonly DAL.Entity dal = new DAL.Entity();
        public Entity()
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
        /// 根据部门代码查找是否存在
        /// </summary>
        public bool ExistsBMDM(string BMDM)
        {

            return dal.ExistsBMDM(BMDM);
        }


          /// <summary>
        /// 根据部门代码查找sjbm
        /// </summary>
        public object ExistsSJBM(string BMDM)
        {
            return dal.ExistsSJBM(BMDM);
        }

        /// <summary>
        /// 增加一条数据
        /// </summary>
        public int Add(Model.Entity model)
        {
            return dal.Add(model);
        }


           /// <summary>
        /// 增加一条数据导入数据
        /// </summary>
        public int AddEnter(Model.Entity model)
        {
            return dal.AddEnter(model);
        
        }

                /// <summary>
        /// 更新一条导入数据
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public int UpdateEnter(Model.Entity model)
        {
            return dal.UpdateEnter(model);
        }

        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.Entity model)
        {
            return dal.Update(model);
        }


        
        /// <summary>
        /// 更新一条数据 部门管理
        /// </summary>
        public int UpdateDepartment(string BMMC, string BMJC, string SJBM, string LXDZ, decimal Lo, decimal La, string BMDM, string FZR, string LXDH, int Sort, string FY, int IsDel, DateTime GXSJ, int ID, string BM, string SJ)
        {
           
            return dal.UpdateDepartment(BMMC, BMJC, SJBM, LXDZ, Lo, La, BMDM, FZR, LXDH, Sort, FY, IsDel, GXSJ, ID, BM, SJ);

        }



        
        /// <summary>
        /// 更新部门是否显示
        /// </summary>
        /// <param name="isdel"></param>
        /// <param name="ID"></param>
        /// <returns></returns>
        //public int UpdateDepartment_Entity_Temp(int isdel, string BMDM)
        //{
        //    return dal.UpdateDepartment_Entity_Temp(isdel, BMDM);
        
        //}

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
        public Model.Entity GetModel(int ID)
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
        ///获取子单位
        /// </summary>
        /// <returns></returns>
        public DataTable GetEntity(string SJBM, int DevType, string BMDM)
        {

            return dal.GetEntity(SJBM, DevType, BMDM);
        }


        /// <summary>
        /// 获取单位终端数
        /// </summary>
        /// <param name="SJBM"></param>
        /// <returns></returns>
        public int GetCount(string SJBM, string BMDM)
        {
            return dal.GetCount(SJBM, BMDM);
        
        }


        /// <summary>
        /// 获取四大队的设备每个类型总数 
        /// </summary>
        /// <returns></returns>
        public int GetAllDevTyCount(int DevType, string s_BMDM)
        {

            return dal.GetAllDevTyCount(DevType, s_BMDM);
        }



        
        /// <summary>
        ///     获取四大队的设备总数
        /// </summary>
        /// <param name="DevType"></param>
        /// <returns></returns>
        public int GetAllCount(string s_BMDM)
        {

            return dal.GetAllCount(s_BMDM);
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
        /// 查找对应BMDM
        /// </summary>
        /// <returns></returns>
        public object GetListName(string BMMC)
        {
            return dal.GetListName(BMMC);
        
        }


             /// <summary>
        /// 根据BMDM查找对应BMDM
        /// </summary>
        /// <returns></returns>
        public object GetListBMMD(string BMDM)
        {
            return dal.GetListBMMD(BMDM);

        }

           /// <summary>
        /// 根据部门名称查找部门编号
        /// </summary>
        /// <returns></returns>
        public object GetListBMDM(string s_BMMC)
        {
            return dal.GetListBMDM(s_BMMC);
        }


        /// <summary>
        /// 获得DropDownList所需字段
        /// </summary>
        /// <returns></returns>
        public DataTable GetList(string s_SJBM, string s_BMDM)
        {

            return dal.GetList(s_SJBM, s_BMDM);

        }


        /// <summary>
        /// 获得DropDownList所需字段
        /// </summary>
        /// <returns></returns>
        public DataTable GetList2(string s_SJBM, string s_BMDM)
        {

            return dal.GetList2(s_SJBM, s_BMDM);

        }

        
          /// <summary>
        /// 获取设备总数量 部门管理
        /// </summary>
        /// <returns></returns>
        public int getcount(string three, string second,string search)
        {
            return dal.getcount(three, second,search);
        }


            /// <summary>
        /// 分页排序设备信息 部门管理
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(string three, string second,string search,int startRowIndex,int maximumRows)
        {
            return dal.getcountpaging(three, second,search, startRowIndex, maximumRows);

        }





        #endregion  BasicMethod

    }
}

