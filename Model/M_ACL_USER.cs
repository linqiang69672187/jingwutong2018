using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model
{
    /// <summary>
    /// ACL_USER:实体类(属性说明自动提取数据库字段的描述信息)
    /// </summary>
    [Serializable]
    public partial class M_ACL_USER
    {
        public M_ACL_USER()
        { }
        #region Model
        private long _id;
        private string _jybh;
        private string _xm;
        private string _bmdm;
        private string _sfzmhm;
        private string _jtzz;
        private int? _jylx;
        private string _mm;
        private int? _jsid;
        private string _sj;
        private string _bgdh;
        private string _bzlx;
        private string _ywgw;
        private string _sgcldj;
        private string _ldjb;
        private DateTime? _csrq;
        private string _xb;
        private string _jx;
        private string _jb;
        private string _jg;
        private DateTime? _rdtsj;
        private string _zzmm;
        private string _mz;
        private string _xl;
        private string _zy;
        private string _zw;
        private DateTime? _rdsj;
        private DateTime? _cgsj;
        private DateTime? _rxzsj;
        private string _zfzgdj;
        private int? _zt;
        private DateTime? _gxsj;
        private DateTime? _cjsj = DateTime.Now;



        public string Statistics()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Append("id:"+ _id+",");
            strSql.Append("jybh:" + _jybh + ",");
            strSql.Append("xm:" + _xm + ",");
            strSql.Append("bmdm:" + _bmdm + ",");
            strSql.Append("sfzmhm:" + _sfzmhm + ",");
            strSql.Append("jtzz:" + _jtzz + ",");
            strSql.Append("jylx:" + _jylx + ",");
            strSql.Append("mm:" + _mm + ",");
            strSql.Append("jsid:" + _jsid + ",");
            strSql.Append("sj:" + _sj + ",");
            strSql.Append("bgdh:" + _bgdh + ",");
            strSql.Append("bzlx:" + _bzlx + ",");
            strSql.Append("ywgw:" + _ywgw + ",");
            strSql.Append("sgcldj:" + _sgcldj + ",");
            strSql.Append("ldjb:" + _ldjb + ",");
            strSql.Append("csrq:" + _csrq + ",");
            strSql.Append("xb:" + _xb + ",");
            strSql.Append("jx:" + _jx + ",");
            strSql.Append("jb:" + _jb + ",");
            strSql.Append("jg:" + _jg + ",");
            strSql.Append("rdtsj:" + _rdtsj + ",");
            strSql.Append("zzmm:" + _zzmm + ",");
            strSql.Append("mz:" + _mz + ",");
            strSql.Append("xl:" + _xl + ",");
            strSql.Append("zy:" + _zy + ",");
            strSql.Append("zw:" + _zw + ",");
            strSql.Append("rdsj:" + _rdsj + ",");
            strSql.Append("cgsj:" + _cgsj + ",");
            strSql.Append("rxzsj:" + _rxzsj + ",");
            strSql.Append("zfzgdj:" + _zfzgdj + ",");
            strSql.Append("zt:" + _zt + ",");
            strSql.Append("gxsj:" + _gxsj + ",");
            strSql.Append("cjsj:" + _cjsj);

                        return strSql.ToString();


        }


        //自加的
        private string _sjbm;
        /// <summary>
        /// 
        /// </summary>
        public long ID
        {
            set { _id = value; }
            get { return _id; }
        }
        /// <summary>
        /// 警员编号
        /// </summary>
        public string JYBH
        {
            set { _jybh = value; }
            get { return _jybh; }
        }
        /// <summary>
        /// 姓名
        /// </summary>
        public string XM
        {
            set { _xm = value; }
            get { return _xm; }
        }
        /// <summary>
        /// 部门代码
        /// </summary>
        public string BMDM
        {
            set { _bmdm = value; }
            get { return _bmdm; }
        }
        /// <summary>
        /// 身份证号码
        /// </summary>
        public string SFZMHM
        {
            set { _sfzmhm = value; }
            get { return _sfzmhm; }
        }
        /// <summary>
        /// 家庭住址
        /// </summary>
        public string JTZZ
        {
            set { _jtzz = value; }
            get { return _jtzz; }
        }
        /// <summary>
        /// 警员类型
        /// </summary>
        public int? JYLX
        {
            set { _jylx = value; }
            get { return _jylx; }
        }
        /// <summary>
        /// 密码
        /// </summary>
        public string MM
        {
            set { _mm = value; }
            get { return _mm; }
        }
        /// <summary>
        /// 警员角色
        /// </summary>
        public int? JSID
        {
            set { _jsid = value; }
            get { return _jsid; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string SJ
        {
            set { _sj = value; }
            get { return _sj; }
        }
        /// <summary>
        /// 
        /// </summary>
        public string BGDH
        {
            set { _bgdh = value; }
            get { return _bgdh; }
        }
        /// <summary>
        /// '编制类型 1-公安编 2-事业编 3-地方编';
        /// </summary>
        public string BZLX
        {
            set { _bzlx = value; }
            get { return _bzlx; }
        }
        /// <summary>
        /// 业务岗位 01-城区管理执勤 02-国省道执勤 03-高速执勤 04-县乡执勤 05-事故处理 06-车驾管 07-道路宣传 08-科技管理 09-其他';
        /// </summary>
        public string YWGW
        {
            set { _ywgw = value; }
            get { return _ywgw; }
        }
        /// <summary>
        /// 事故处理等级  0	无等级1	高级2	中级3	初级';
        /// </summary>
        public string SGCLDJ
        {
            set { _sgcldj = value; }
            get { return _sgcldj; }
        }
        /// <summary>
        /// 领导级别 D0	总队领导D1	支队领导D2	大队领导D3	中队领导ZZ	其他'
        /// </summary>
        public string LDJB
        {
            set { _ldjb = value; }
            get { return _ldjb; }
        }
        /// <summary>
        /// 出生日期
        /// </summary>
        public DateTime? CSRQ
        {
            set { _csrq = value; }
            get { return _csrq; }
        }
        /// <summary>
        /// 性别
        /// </summary>
        public string XB
        {
            set { _xb = value; }
            get { return _xb; }
        }
        /// <summary>
        /// 警衔
        /// </summary>
        public string JX
        {
            set { _jx = value; }
            get { return _jx; }
        }
        /// <summary>
        /// 职级
        /// </summary>
        public string JB
        {
            set { _jb = value; }
            get { return _jb; }
        }
        /// <summary>
        /// 籍贯
        /// </summary>
        public string JG
        {
            set { _jg = value; }
            get { return _jg; }
        }
        /// <summary>
        /// 入党团时间
        /// </summary>
        public DateTime? RDTSJ
        {
            set { _rdtsj = value; }
            get { return _rdtsj; }
        }
        /// <summary>
        /// 政治面貌
        /// </summary>
        public string ZZMM
        {
            set { _zzmm = value; }
            get { return _zzmm; }
        }
        /// <summary>
        /// 民族
        /// </summary>
        public string MZ
        {
            set { _mz = value; }
            get { return _mz; }
        }
        /// <summary>
        /// 学历
        /// </summary>
        public string XL
        {
            set { _xl = value; }
            get { return _xl; }
        }
        /// <summary>
        /// 所学专业
        /// </summary>
        public string ZY
        {
            set { _zy = value; }
            get { return _zy; }
        }
        /// <summary>
        /// 岗位职务
        /// </summary>
        public string ZW
        {
            set { _zw = value; }
            get { return _zw; }
        }
        /// <summary>
        /// 入队时间
        /// </summary>
        public DateTime? RDSJ
        {
            set { _rdsj = value; }
            get { return _rdsj; }
        }
        /// <summary>
        /// 参加工作时间
        /// </summary>
        public DateTime? CGSJ
        {
            set { _cgsj = value; }
            get { return _cgsj; }
        }
        /// <summary>
        /// 任现时间
        /// </summary>
        public DateTime? RXZSJ
        {
            set { _rxzsj = value; }
            get { return _rxzsj; }
        }
        /// <summary>
        /// 执法资格等级
        /// </summary>
        public string ZFZGDJ
        {
            set { _zfzgdj = value; }
            get { return _zfzgdj; }
        }
        /// <summary>
        /// 警员状态
        /// </summary>
        public int? ZT
        {
            set { _zt = value; }
            get { return _zt; }
        }
        /// <summary>
        /// 更新时间
        /// </summary>
        public DateTime? GXSJ
        {
            set { _gxsj = value; }
            get { return _gxsj; }
        }
        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime? CJSJ
        {
            set { _cjsj = value; }
            get { return _cjsj; }
        }

        /// <summary>
        /// 自加的 上级部门
        /// </summary>
        public string SJBM
        {
            set { _sjbm = value; }
            get { return _sjbm; }
        }
        #endregion Model

    }
}
