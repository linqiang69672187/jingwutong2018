using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// Device:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class Device
    {
        public Device()
        { }
        #region Model
        private long _id;
        private string _bmdm;
        private int _devtype;
        private string _devid;
        private string _cartype;
        private string _platenumber;
        private int? _workstate;
        private int? _allocatestate;
        private DateTime? _allocatetime;
        private string _jybh;
        private string _imei;
        private string _simid;
        private string _manufacturer;
        private string _sbxh;
        private string _sbgg;
        private string _projname;
        private string _projnum;
        private string _price;
        private string _xmfzr;
        private string _xmfzrdh;
        private DateTime? _cgsj;
        private DateTime? _bxq;
        private DateTime? _bfqx;
        private DateTime? _bcjysj;
        private DateTime? _xcjysj;
        private DateTime? _creatdate = DateTime.Now;

        public string Statistics()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Append("id:" + _id + ",");
            strSql.Append("bmdm:" + _bmdm + ",");
            strSql.Append("devtype:" + _devtype + ",");


            strSql.Append("devid:" + _devid + ",");
            strSql.Append("cartype:" + _cartype + ",");
            strSql.Append("platenumber:" + _platenumber + ",");

            strSql.Append("workstate:" + _workstate + ",");
            strSql.Append("allocatestate:" + _allocatestate + ",");


            strSql.Append("allocatetime:" + _allocatetime + ",");
            strSql.Append("jybh:" + _jybh + ",");
            strSql.Append("imei:" + _imei + ",");
            strSql.Append("simid:" + _simid + ",");
            strSql.Append("manufacturer:" + _manufacturer + ",");
            strSql.Append("sbxh:" + _sbxh + ",");
            strSql.Append("sbgg:" + _sbgg + ",");
            strSql.Append("projname:" + _projname + ",");
            strSql.Append("projnum:" + _projnum + ",");
            strSql.Append("price:" + _price + ",");
            strSql.Append("xmfzr:" + _xmfzr + ",");
            strSql.Append("xmfzrdh :" + _xmfzrdh + ",");
            strSql.Append("cgsj:" + _cgsj + ",");
            strSql.Append("bxq:" + _bxq + ",");
            strSql.Append("bfq:" + _bfqx + ",");
            strSql.Append("bcjysj:" + _bcjysj + ",");
            strSql.Append("xcjysj :" + _xcjysj + ",");
            strSql.Append("creatdate:" + _creatdate + ",");

            return strSql.ToString();
        
        
        }




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
        public string BMDM
        {
            set { _bmdm = value; }
            get { return _bmdm; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int DevType
        {
            set { _devtype = value; }
            get { return _devtype; }
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
        public string Cartype
        {
            set { _cartype = value; }
            get { return _cartype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string PlateNumber
        {
            set { _platenumber = value; }
            get { return _platenumber; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? WorkState
        {
            set { _workstate = value; }
            get { return _workstate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? AllocateState
        {
            set { _allocatestate = value; }
            get { return _allocatestate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? AllocateTime
        {
            set { _allocatetime = value; }
            get { return _allocatetime; }
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
        public string IMEI
        {
            set { _imei = value; }
            get { return _imei; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string SIMID
        {
            set { _simid = value; }
            get { return _simid; }
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
        public string ProjName
        {
            set { _projname = value; }
            get { return _projname; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string ProjNum
        {
            set { _projnum = value; }
            get { return _projnum; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string Price
        {
            set { _price = value; }
            get { return _price; }
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
        /// <summary>
        /// 
        /// </summary>
        public DateTime? CreatDate
        {
            set { _creatdate = value; }
            get { return _creatdate; }
        }
        #endregion Model

    }
}
