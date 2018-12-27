using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// Alarm_EveryDayInfo:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class M_Alarm_EveryDayInfo
    {
        public M_Alarm_EveryDayInfo()
        { }
        #region Model
        private long _id;
        private string _devid;
        private DateTime? _alarmday;
        private int? _alarmstate;
        private int? _alarmtype;
        private string _contacts;
        private string _entity;
        private string _tel;
        private int? _devtype;
        private int? _value;
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
        public string DevId
        {
            set { _devid = value; }
            get { return _devid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? AlarmDay
        {
            set { _alarmday = value; }
            get { return _alarmday; }
        }
        /// <summary>
        /// "1"表示告警，“0”表示正常
        /// </summary>
        public int? AlarmState
        {
            set { _alarmstate = value; }
            get { return _alarmstate; }
        }
        /// <summary>
        /// “1”表示在线时长，“2”表示处理量
        /// </summary>
        public int? AlarmType
        {
            set { _alarmtype = value; }
            get { return _alarmtype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string Contacts
        {
            set { _contacts = value; }
            get { return _contacts; }
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
        public int? DevType
        {
            set { _devtype = value; }
            get { return _devtype; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? Value
        {
            set { _value = value; }
            get { return _value; }
        }
        #endregion Model

    }
}
