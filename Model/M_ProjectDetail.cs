using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// ProjectDetail:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class M_ProjectDetail
    {
        public M_ProjectDetail()
        { }
        #region Model
        private int _id;
        private string _xmbh;
        private string _sbxh;
        private string _sbgg;
        private int? _price;
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
        public string SBXH
        {
            set { _sbxh = value; }
            get { return _sbxh; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string SBGG
        {
            set { _sbgg = value; }
            get { return _sbgg; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? Price
        {
            set { _price = value; }
            get { return _price; }
        }
        #endregion Model

    }
}
