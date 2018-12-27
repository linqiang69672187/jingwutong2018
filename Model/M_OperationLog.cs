using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// OperationLog:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class M_OperationLog
    {
        public M_OperationLog()
        { }
        #region Model
        private long _id;
        private string _jybh;
        private string _module;
        private string _opercontent;
        private string _ipaddr;
        private DateTime? _logtime;
        private string _bz;
        private string _optobject;

        /// <summary>
        /// 
        /// </summary>
        public long ID
        {
            set { _id = value; }
            get { return _id; }
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
        public string Module
        {
            set { _module = value; }
            get { return _module; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string OperContent
        {
            set { _opercontent = value; }
            get { return _opercontent; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string IpAddr
        {
            set { _ipaddr = value; }
            get { return _ipaddr; }
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
        public string BZ
        {
            set { _bz = value; }
            get { return _bz; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string OptObject
        {
            set { _optobject = value; }
            get { return _optobject; }
        }

        #endregion Model

    }
}
