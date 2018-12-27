using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{

    /// <summary>
    /// ProjectInfo:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class M_ProjectInfo
    {
        public M_ProjectInfo()
        { }
        #region Model
        private int _id;
        private string _xmbh;
        private string _xmmc;
        private string _xmfzr;
        private string _xmfzrdh;
        private DateTime? _cgsj;
        private DateTime? _bxq;
        private DateTime? _bfqx;
        private string _manufacturer;
        private DateTime? _bcjysj;
        private DateTime? _xcjysj;
        /// <summary>
        /// 
        /// </summary>
        public int ID
        {
            set { _id = value; }
            get { return _id; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string XMBH
        {
            set { _xmbh = value; }
            get { return _xmbh; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string XMMC
        {
            set { _xmmc = value; }
            get { return _xmmc; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string XMFZR
        {
            set { _xmfzr = value; }
            get { return _xmfzr; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string XMFZRDH
        {
            set { _xmfzrdh = value; }
            get { return _xmfzrdh; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? CGSJ
        {
            set { _cgsj = value; }
            get { return _cgsj; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? BXQ
        {
            set { _bxq = value; }
            get { return _bxq; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? BFQX
        {
            set { _bfqx = value; }
            get { return _bfqx; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string Manufacturer
        {
            set { _manufacturer = value; }
            get { return _manufacturer; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? BCJYSJ
        {
            set { _bcjysj = value; }
            get { return _bcjysj; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? XCJYSJ
        {
            set { _xcjysj = value; }
            get { return _xcjysj; }
        }
        #endregion Model

    }
}
