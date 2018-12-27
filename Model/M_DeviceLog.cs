using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// DeviceLog:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class M_DeviceLog
    {
        public M_DeviceLog()
        { }
        #region Model
        private long? _id;
        private int? _logtype;
        private string _devid;
        private int? _devtype;
        private int? _devstate;
        private string _username;
        private string _entity;
        private string _tel;
        private string _jybh;
        private DateTime? _logtime;
        private string _bxr;
        private string _bz;
        private int? _allocatestate;


        public string Statistics()
        {
            StringBuilder strSql = new StringBuilder();

        strSql.Append("id:"+ _id+",");
        strSql.Append("logtype:" + _logtype + ",");
        strSql.Append("devid:" + _devid + ",");
        strSql.Append("devtype:" + _devtype + ",");
        strSql.Append("devstate:" + _devstate + ",");
        strSql.Append("username:" + _username + ",");
        strSql.Append("entity:" + _entity + ",");
        strSql.Append("tel:" + _tel + ",");
        strSql.Append("jybh:" + _jybh + ",");
        strSql.Append("logtime:" + _logtime + ",");
        strSql.Append("bxr:" + _bxr + ",");
        strSql.Append("bz:" + _bz + ",");
        strSql.Append("allocatestate:" + _allocatestate );

        return strSql.ToString();

        }


        /// <summary>
        /// 
        /// </summary>
        public long? ID
        {
            set { _id = value; }
            get { return _id; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? LogType
        {
            set { _logtype = value; }
            get { return _logtype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string DevId
        {
            set { _devid = value; }
            get { return _devid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? DevType
        {
            set { _devtype = value; }
            get { return _devtype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? DevState
        {
            set { _devstate = value; }
            get { return _devstate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string UserName
        {
            set { _username = value; }
            get { return _username; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string Entity
        {
            set { _entity = value; }
            get { return _entity; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string Tel
        {
            set { _tel = value; }
            get { return _tel; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string JYBH
        {
            set { _jybh = value; }
            get { return _jybh; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? LogTime
        {
            set { _logtime = value; }
            get { return _logtime; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string BXR
        {
            set { _bxr = value; }
            get { return _bxr; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string BZ
        {
            set { _bz = value; }
            get { return _bz; }
        }


        /// <summary>
        /// 
        /// </summary>
        public int? AllocateState
        {
            set { _allocatestate = value; }
            get { return _allocatestate; }
        }
        #endregion Model

    }
}
