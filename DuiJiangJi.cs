using System;
using System.Collections.Generic;
using System.Text;

namespace JingWuTong
{
    public class DuiJiangJi
    {

        public DuiJiangJi()
        {

        }

        public static string ByteToString(byte[] InBytes)
        {
            string StringOut = "";
            foreach (byte InByte in InBytes)
            {
                StringOut = StringOut + String.Format("{0:X2} ", InByte);
            }
            return StringOut;
        }


        public static byte[] HexStrTobyte(string hexString)
        {
            hexString = hexString.Replace(" ", "");
            if ((hexString.Length % 2) != 0)
                hexString += " ";
            byte[] returnBytes = new byte[hexString.Length / 2];
            for (int i = 0; i < returnBytes.Length; i++)
                returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2).Trim(), 16);
            return returnBytes;
        }


        public static Decimal ChangeToDecimal(string strData)
        {
            Decimal dData = 0.0M;
            if (strData.Contains("E"))
            {
                dData = Convert.ToDecimal(Decimal.Parse(strData.ToString(), System.Globalization.NumberStyles.Float));
            }
            else
            {
                dData = Convert.ToDecimal(strData);
            }
            return Math.Round(dData, 4);
        }

        public static Byte[] GetASCIIBytes(string str)
        {
            return Encoding.ASCII.GetBytes(str);
            //for (var i=0;i<bytes.Length;i++)
            //{
            //    bytes[i] = Convert.ToByte(string.Format("{0:X}",bytes[i]));
            //}
            //return bytes;
        }

        /// <summary>
        /// 网络字节序
        /// </summary>
        /// <param name="bytes"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="value"></param>
        public static void Converts(ref Byte[] bytes, int start, int end, Double value)
        {
            var tmpbytes = BitConverter.GetBytes(value);
            Array.Reverse(tmpbytes);
            for (var i = start; i <= end; i++)
            {
                bytes[i] = tmpbytes[i - start];
            }
        }

        public static byte[] getBbytes(Byte[] bytes, int start, int length)
        {
            byte[] bb = new byte[length];
            int k = 0;

            for (var i = start; i < start + length; i++)
            {
                bb[k] = bytes[i];
                k++;
            }
            return bb;
        }
    }
}
