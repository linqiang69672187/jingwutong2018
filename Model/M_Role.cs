using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// Role:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class M_Role
    {
        public M_Role()
        { }
        #region Model
        private int _id;
        private string _rolename;
        private string _power;
        private int? _status;
        private string _bz;
        private string _crateator;
        private DateTime? _createdate = DateTime.Now;


        public string Statistics()
        {
            StringBuilder strSql = new StringBuilder();
              strSql.Append("id"+_id+",");
              strSql.Append("id" + _rolename+",");
              strSql.Append("id" + _power+",");
              strSql.Append("id" + _status+",");
              strSql.Append("id" + _bz+",");
              strSql.Append("id" + _crateator+",");
              strSql.Append("id" + _createdate);


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
        public string RoleName
        {
            set { _rolename = value; }
            get { return _rolename; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string Power
        {
            set { _power = value; }
            get { return _power; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int? Status
        {
            set { _status = value; }
            get { return _status; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string Bz
        {
            set { _bz = value; }
            get { return _bz; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string Crateator
        {
            set { _crateator = value; }
            get { return _crateator; }
        }
        /// <summary>
        /// 
        /// </summary>
        public DateTime? CreateDate
        {
            set { _createdate = value; }
            get { return _createdate; }
        }
        #endregion Model

    }
}
