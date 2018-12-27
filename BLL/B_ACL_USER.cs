using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    /// <summary>
    /// ACL_USER
    /// </summary>
    public partial class B_ACL_USER
    {
        private readonly DAL.D_ACL_USER dal = new DAL.D_ACL_USER();
        public B_ACL_USER()
        { }
        #region  BasicMethod
        /// <summary>
        /// 是否存在该记录
        /// </summary>
        public object Exists(string JYBH, string password)
        {
            return dal.Exists(JYBH, password);
        }



                /// <summary>
        /// 用户登录查找姓名
        /// </summary>
        /// <param name="JYBH"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public object ExistsXM(string JYBH, string password)
        {

            return dal.ExistsXM(JYBH, password);
        }


                /// <summary>
        /// 根据警员编号查找姓名
        /// </summary>
        /// <param name="JYBH"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public object ExistsJYXM(string JYBH)
        {

            return dal.ExistsJYXM(JYBH);
        }


               /// <summary>
        /// 得到一部分对象实体 用于登录操作
        /// </summary>
        /// <param name="ID"></param>
        /// <returns></returns>
        public Model.M_ACL_USER ExistsModel(string JYBH, string password)
        {
            return dal.ExistsModel(JYBH, password);

        }


                /// <summary>
        /// 查找警员编号是否重复
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public bool ExistsJYBH(string JYBH)
        {

            return dal.ExistsJYBH(JYBH);
        }

        /// <summary>
        /// 增加一部分数据
        /// </summary>
        public int AddLittle(Model.M_ACL_USER model)
        {
            return dal.AddLittle(model);
        }


        
        /// <summary>
        /// 警员表新增
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public int AddACL_USER(Model.M_ACL_USER model)
        {

            return dal.AddACL_USER(model);
        }


        /// <summary>
        /// 增加一条导入数据
        /// </summary>
        public int Add(Model.M_ACL_USER model)
        {
            return dal.Add(model);
        }


           /// <summary>
        /// 更新一条导入数据
        /// </summary>
        public int UpdateData(Model.M_ACL_USER model)
        {
            return dal.UpdateData(model);
        }
        /// <summary>
        /// 更新一条部分数据
        /// </summary>
        public int UpdatePart(Model.M_ACL_USER model) 
        {

            return dal.UpdatePart(model);
        
        }



        
        /// <summary>
        ///警员表更新
        /// </summary>
        public int UpdateACL_USE(Model.M_ACL_USER model)
        {
            return dal.UpdateACL_USE(model);
        }


        /// <summary>
        /// 更新一条数据
        /// </summary>
        public int Update(Model.M_ACL_USER model)
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
        /// 删除一条数据根据警员编号来
        /// </summary>
        public bool DeleteJYBH(string JYBH)
        {
            return dal.DeleteJYBH(JYBH);
        
        }



        /// <summary>
        /// 得到一部分对象实体
        /// </summary>
        public Model.M_ACL_USER GetLittleModel(string JYBH)
        {

            return dal.GetLittleModel(JYBH);
        }




        /// <summary>
        /// 得到一个对象实体
        /// </summary>
        public Model.M_ACL_USER GetModel(long ID)
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
        /// 更具 角色ID 查找角色是否存在 JSID
        /// </summary>
        /// <returns></returns>
        public bool ExistsJSID(int JSID)
        {
            return dal.ExistsJSID(JSID);
        
        }



        /// <summary>
        /// 获取设备总数量
        /// </summary>
        /// <returns></returns>
        public int getcount(string three, string second, int JYLX, int JSID, int LDJB, string search)
        {
            return dal.getcount(three, second, JYLX, JSID, LDJB, search);

        }


        /// <summary>
        /// 分页排序设备信息
        /// </summary>
        /// <returns></returns>
        public DataTable getcountpaging(string three, string second, int JYLX, int JSID, int LDJB, string search, int startRowIndex, int maximumRows)
        {

            return dal.getcountpaging(three, second, JYLX, JSID, LDJB, search, startRowIndex, maximumRows);
        }



        #endregion  BasicMethod

    }
}

