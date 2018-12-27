using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// AlarmGate:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class M_AlarmGate
    {
        public M_AlarmGate()
        { }
        #region Model
        private int _id;
        private int? _devtype;
        private int? _alarmtype;
        private int? _commonalarmgate;
        private int? _urgencyalarmgate;
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
        public int? DevType
        {
            set { _devtype = value; }
            get { return _devtype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? AlarmType
        {
            set { _alarmtype = value; }
            get { return _alarmtype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? CommonAlarmGate
        {
            set { _commonalarmgate = value; }
            get { return _commonalarmgate; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? UrgencyAlarmGate
        {
            set { _urgencyalarmgate = value; }
            get { return _urgencyalarmgate; }
        }
        #endregion Model

    }
}
