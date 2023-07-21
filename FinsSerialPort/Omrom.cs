using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;

namespace Measurement.FinsSerialPort
{
    internal class Omrom
    {
        // Fins 头 节点 固定指令 等待时间 
        private const string Head = "@00FA0";

        // ICF DA2 SA2 SID 
        private const string Mark = "00000000";
        private const string HeadMark = "@00FA000000000";
        private readonly SerialPort serialPort = new SerialPort();

        // 初始化
        public void Init(string port)
        {
            serialPort.BaudRate = 115200;
            serialPort.DataBits = 8;
            serialPort.StopBits = StopBits.One;
            serialPort.Parity = Parity.None;
            serialPort.ReadBufferSize = 4096;
            serialPort.PortName = port;
        }

        // 操作指令 
        // IO 读
        private const string IoRead = "0101";

        // IO 写
        private const string IoWrite = "0102";

        // 存储器
        // Cio BIT
        private const string Cio = "30";

        // W BIT
        private const string W = "31";

        // WORD 地址
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
        public void SetPort(string s)
        {
            serialPort.PortName = s;
        }

        // 串口名获取
        public string GetPort()
        {
            return serialPort.PortName;
        }

        // 串口列表
        public string[] GetPortList()
        {
            return SerialPort.GetPortNames().OrderBy(SortPort).ToArray();
        }

        // 串口数组排序
        private int SortPort(string port)
        {
            // 判断是否为COM开头并且解析字符串3位后的数字
            if (port.StartsWith("COM", StringComparison.OrdinalIgnoreCase) && int.TryParse(port.Substring(3), out var n))

                // 是数字则返回数字
                return n;

            // 否则返回0
            return 0;
        }

        // 串口发送
        private bool SeriaWrite(byte[] chars, int v, int count)
        {
            try
            {
                serialPort.Open();
                serialPort.Write(chars, v, count);
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }

        //校验
        private static string Fcs(string s)
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
            Thread.Sleep(200);

            // 获取缓冲个数
            var n = serialPort.BytesToRead;
            var str = new byte[n];

            //读取
            serialPort.Read(str, 0, n);
            var result = new List<string>();

            //转换
            foreach (var b in str) result.Add(((char)b).ToString());

            serialPort.Close();
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