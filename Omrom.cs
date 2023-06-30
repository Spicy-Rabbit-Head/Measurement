using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Text;
using System.Threading;

namespace Measurement
{
    internal class Omrom
    {
        // Fins 头 节点 固定指令 等待时间 
        private readonly string head = "@00FA0";

        // ICF DA2 SA2 SID 
        private readonly string mark = "00000000";
        private const string HeadMark = "@00FA000000000";
        private static readonly SerialPort SerialPort = new SerialPort();

        public void Init()
        {
            SerialPort.BaudRate = 115200;
            SerialPort.DataBits = 8;
            SerialPort.StopBits = StopBits.One;
            SerialPort.Parity = Parity.None;
            SerialPort.ReadBufferSize = 4096;
        }

        // 操作指令 
        // IO 读
        private const string IoRead = "0101";

        // IO 写
        private const string IoWrite = "0102";

        // 存储器
        // Cio Bit
        private const string Cio = "30";

        // W Bit
        private const string W = "31";

        // Word 地址
        // 上
        private const string Word1 = "000008";

        // 下
        private const string Word2 = "000004";

        // 气缸上
        private const string Word3 = "000007";

        // 气缸下
        private const string Word4 = "00000A";

        // bit 地址 
        private const string Bit = "0001";

        // 状态
        private const string On = "01";
        private const string Off = "00";

        // 结束
        private const string Cr = "*";

        //串口名设置
        public static void SetPort(string s)
        {
            SerialPort.PortName = s;
        }

        // 串口发送
        private bool SeriaWrite(byte[] chars, int v, int count)
        {
            try
            {
                SerialPort.Open();
                SerialPort.Write(chars, v, count);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        //校验
        private string Fcs(string s)
        {
            // 提取字节数组
            var b = Encoding.ASCII.GetBytes(s);
            // xor 存放校验结果 初值去首元素值
            var xorResult = b[0];
            // 求xor校验和 从第二位开始
            for (var i = 1; i < b.Length; i++)
                //进行异或运算
                xorResult ^= b[i];

            // 返回 ASCII 异或值
            return Convert.ToString(xorResult, 16).ToUpper();
        }

        //转换16进制并发送
        private bool Conversion(string s)
        {
            var fcs = Fcs(s);
            if (fcs.Length == 1) s += '0';

            s += fcs;
            s += Cr;
            var send = new List<byte>(Encoding.ASCII.GetBytes(s));
            send.Add(0x0D);
            return SeriaWrite(send.ToArray(), 0, send.Count);
        }

        // 数据接收
        private List<string> DataReceived()
        {
            Thread.Sleep(300);
            // 获取缓冲个数
            var n = SerialPort.BytesToRead;
            var str = new byte[n];
            //读取
            SerialPort.Read(str, 0, n);
            var result = new List<string>();
            //转换
            foreach (var b in str) result.Add(((char)b).ToString());

            SerialPort.Close();
            return result;
        }


        // 上丝杆动作
        public bool UpperScrewAction()
        {
            var str = HeadMark + IoWrite + Cio + Word1 + Bit + On;
            if (Conversion(str))
            {
                var result = DataReceived();
                //if (result[24] == "0")
                //{
                //    return true;
                //}
                //return false;
                return true;
            }

            return false;
            // @00FA00400000000102000040*\CR
        }

        //下丝杆动作 
        public bool LowerScrewAction()
        {
            var str = HeadMark + IoWrite + W + Word2 + Bit + On;
            if (Conversion(str))
            {
                var result = DataReceived();
                //if (result[24] == "0")
                //{
                //    return true;
                //}
                //return false;
                return true;
            }

            return false;
            // @00FA00400000000102000040*\CR
        }

        // 气缸上状态
        public bool CylinderTop()
        {
            var str = HeadMark + IoRead + Cio + Word3 + Bit;
            if (Conversion(str))
            {
                var result = DataReceived();
                //if (result[24] == "1")
                //{
                //    return true;
                //}
                //return false;
                return true;
            }

            return false;
            // on  @00FA0040000000010100000142*\CR
            // off @00FA0040000000010100000043*\CR
        }

        // 气缸下状态
        public bool CylinderButton()
        {
            var str = HeadMark + IoRead + Cio + Word4 + Bit;
            if (Conversion(str))
            {
                var result = DataReceived();
                //   if (result[24] == "1")
                //{
                //    return true;
                //}
                //return false;
                return true;
            }

            return false;
            // on  @00FA0040000000010100000142*\CR
            // off @00FA0040000000010100000043*\CR
        }
    }
}