using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// Entity:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class Entity
    {
        public Entity()
        { }
        #region Model
        private int _id;
        private string _bmmc;
        private string _bmqc;
        private string _bmjc;
        private string _bmdm;
        private string _sjbm;
        private string _depth;
        private string _picurl;
        private int? _usercount;
        private string _yzmc;
        private string _fzjg;
        private string _bmjb;
        private string _kclyw;
        private string _ywlb;
        private string _czhm;
        private string _lxdz;
        private string _txzqfr;
        private string _fzr;
        private string _lxr;
        private string _lxdh;
        private string _jkyh;
        private string _jzjb;
        private string _gltz;
        private string _jfly;
        private string _yfly;
        private string _jlzt = "1";
        private string _jrgaw = "1";
        private string _lsgx = "2";
        private string _jzptglbm;
        private string _csbj = "0";
        private string _fy;
        private string _fyjg;
        private int? _sort;
        private decimal? _lo;
        private decimal? _la;
        private string _bz;
        private DateTime? _cjsj = DateTime.Now;
        private DateTime? _gxsj;
        private int? _isdel = 0;

        public string Statistics()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Append("id:" + _id + ",");
            strSql.Append("bmmc:" + _bmmc + ",");
            strSql.Append("bmqc:" + _bmqc + ",");
            strSql.Append("bmjc:" + _bmjc + ",");
            strSql.Append("bmdm:" + _bmdm + ",");
            strSql.Append("sjbm:" + _sjbm + ",");
            strSql.Append("depth:" + _depth + ",");
            strSql.Append("picurl:" + _picurl + ",");
            strSql.Append("usercount:" + _usercount + ",");
            strSql.Append("yzmc:" + _yzmc + ",");
            strSql.Append("fzjg:" + _fzjg + ",");
            strSql.Append("bmjb:" + _bmjb + ",");
            strSql.Append("kclyw:" + _kclyw + ",");
            strSql.Append("ywlb:" + _ywlb + ",");
            strSql.Append("czhm:" + _czhm + ",");
            strSql.Append("lxdz:" + _lxdz + ",");
            strSql.Append("txzqfr:" + _txzqfr + ",");
            strSql.Append("fzr:" + _fzr + ",");
            strSql.Append("lxr:" + _lxr + ",");
            strSql.Append("lxdh:" + _lxdh + ",");
            strSql.Append("jkyh:" + _jkyh + ",");
            strSql.Append("jzjb:" + _jzjb + ",");
            strSql.Append("gltz:" + _gltz + ",");
            strSql.Append("jfly:" + _jfly + ",");
            strSql.Append("yfly:" + _yfly + ",");
            strSql.Append("jlzt:" + "1" + ",");
            strSql.Append("jrgaw:" + "1" + ",");
            strSql.Append("lsgx:" + "2" + ",");
            strSql.Append("jzptglbm:" + _jzptglbm + ",");
            strSql.Append("csbj:" + "0" + ",");
            strSql.Append("fy:" + _fy + ",");
            strSql.Append("fyjg:" + _fyjg + ",");
            strSql.Append("sor:" + _sort + ",");
            strSql.Append("lo:" + _lo + ",");
            strSql.Append("la:" + _la + ",");
            strSql.Append("bz:" + _bz + ",");
            strSql.Append("cjsj:" + _cjsj + ",");
            strSql.Append("gxsj" + _gxsj + ",");
             return strSql.ToString();
        }


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
        public string BMMC
        {
            set { _bmmc = value; }
            get { return _bmmc; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string BMQC
        {
            set { _bmqc = value; }
            get { return _bmqc; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string BMJC
        {
            set { _bmjc = value; }
            get { return _bmjc; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string BMDM
        {
            set { _bmdm = value; }
            get { return _bmdm; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string SJBM
        {
            set { _sjbm = value; }
            get { return _sjbm; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string Depth
        {
            set { _depth = value; }
            get { return _depth; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string PicUrl
        {
            set { _picurl = value; }
            get { return _picurl; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? UserCount
        {
            set { _usercount = value; }
            get { return _usercount; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string YZMC
        {
            set { _yzmc = value; }
            get { return _yzmc; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FZJG
        {
            set { _fzjg = value; }
            get { return _fzjg; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string BMJB
        {
            set { _bmjb = value; }
            get { return _bmjb; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string KCLYW
        {
            set { _kclyw = value; }
            get { return _kclyw; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string YWLB
        {
            set { _ywlb = value; }
            get { return _ywlb; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string CZHM
        {
            set { _czhm = value; }
            get { return _czhm; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string LXDZ
        {
            set { _lxdz = value; }
            get { return _lxdz; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string TXZQFR
        {
            set { _txzqfr = value; }
            get { return _txzqfr; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string FZR
        {
            set { _fzr = value; }
            get { return _fzr; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string LXR
        {
            set { _lxr = value; }
            get { return _lxr; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string LXDH
        {
            set { _lxdh = value; }
            get { return _lxdh; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string JKYH
        {
            set { _jkyh = value; }
            get { return _jkyh; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string JZJB
        {
            set { _jzjb = value; }
            get { return _jzjb; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string GLTZ
        {
            set { _gltz = value; }
            get { return _gltz; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string JFLY
        {
            set { _jfly = value; }
            get { return _jfly; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string YFLY
        {
            set { _yfly = value; }
            get { return _yfly; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string JLZT
        {
            set { _jlzt = value; }
            get { return _jlzt; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string JRGAW
        {
            set { _jrgaw = value; }
            get { return _jrgaw; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string LSGX
        {
            set { _lsgx = value; }
            get { return _lsgx; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string JZPTGLBM
        {
            set { _jzptglbm = value; }
            get { return _jzptglbm; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string CSBJ
        {
            set { _csbj = value; }
            get { return _csbj; }
        }
        /// <summary>
        /// 法院
        /// </summary>
        public string FY
        {
            set { _fy = value; }
            get { return _fy; }
        }
        /// <summary>
        /// 复议机构
        /// </summary>
        public string FYJG
        {
            set { _fyjg = value; }
            get { return _fyjg; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? Sort
        {
            set { _sort = value; }
            get { return _sort; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal? Lo
        {
            set { _lo = value; }
            get { return _lo; }
        }
        /// <summary>
        /// 
        /// </summary>
        public decimal? La
        {
            set { _la = value; }
            get { return _la; }
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
        public DateTime? CJSJ
        {
            set { _cjsj = value; }
            get { return _cjsj; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? GXSJ
        {
            set { _gxsj = value; }
            get { return _gxsj; }
        }

        /// <summary>
        /// 1：正常；2：删除；
        /// </summary>
        public int? IsDel
        {
            set { _isdel = value; }
            get { return _isdel; }
        }
        #endregion Model

    }
}
