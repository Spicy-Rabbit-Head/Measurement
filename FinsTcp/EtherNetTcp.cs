using System;
using System.Net.Sockets;
using System.Threading;

namespace Measurement.FinsTcp
{
    public class EtherNetTcp
    {
        private readonly TcpClient client;
        private NetworkStream stream;
        private Timer timer;

        /// <summary>
        /// PLC节点号
        /// </summary>
        private byte plcNode { get; set; }

        /// <summary>
        /// PC节点号
        /// </summary>
        private byte pcNode { get; set; }

        /// <summary>
        /// 发送数据到远程服务器。
        /// </summary>
        /// <param name="sd">要发送的字节数组</param>
        /// <returns>发送操作结果：
        ///     0 表示发送成功；
        ///    -1 表示发送失败或没有连接
        /// </returns>
        private short SendData(byte[] sd)
        {
            // 检查是否连接
            if (stream == null)
                return -1;
            try
            {
                // 向服务器发送数据
                stream.Write(sd, 0, sd.Length);
                return 0;
            }
            catch
            {
                // 发送失败或发送异常
                return -1;
            }
        }

        /// <summary>
        /// 从远程服务器接收数据。
        /// </summary>
        /// <param name="rd">用于接收数据的字节数组</param>
        /// <returns>接收操作结果：
        ///     0 表示接收成功；
        ///    -1 表示接收失败或没有连接
        /// </returns>
        private short ReceiveData(byte[] rd)
        {
            // 检查是否有连接
            if (stream == null)
                return -1;

            try
            {
                // 等待可读数据，直到接收到指定长度的数据
                var index = 0;
                do
                {
                    var len = stream.Read(rd, index, rd.Length - index);

                    // 读取不到数据时就跳出,网络异常断开，数据读取不完整
                    if (len == 0) return -1;

                    // 读取到数据时，累加已读取的数据长度
                    index += len;
                } while (index < rd.Length);

                // 接收成功
                return 0;
            }
            catch
            {
                // 接收失败或接收异常
                return -1;
            }
        }

        /// <summary>
        /// Fins读写指令生成
        /// </summary>
        /// <param name="rw">读写类型</param>
        /// <param name="mr">寄存器类型</param>
        /// <param name="mt">地址类型</param>
        /// <param name="ch">起始地址</param>
        /// <param name="offset">位地址：00-15,字地址则为00</param>
        /// <param name="cnt">地址个数,按位读写只能是1</param>
        /// <returns></returns>
        private byte[] FinsCmd(RorW rw, PlcMemory mr, MemoryType mt, short ch, short offset, short cnt)
        {
            // 用于存储 FINS 指令的字节数组，长度为 34
            var array = new byte[34];

            // TCP FINS 头
            // F
            array[0] = 0x46;

            // I
            array[1] = 0x49;

            // N
            array[2] = 0x4E;

            // S
            array[3] = 0x53;

            // cmd 长度
            array[4] = 0;
            array[5] = 0;

            // 指令长度从下面字节开始计算array[8]
            if (rw == RorW.READ)
            {
                array[6] = 0;

                // 26 读指令长度为 26 字节
                array[7] = 0x1A;
            }
            else
            {
                // 写数据的时候一个字占两个字节，而一个位只占一个字节
                if (mt == MemoryType.WORD)
                {
                    array[6] = (byte)((cnt * 2 + 26) / 256);
                    array[7] = (byte)((cnt * 2 + 26) % 256);
                }
                else
                {
                    array[6] = 0;

                    // 27 写指令长度为 27 字节
                    array[7] = 0x1B;
                }
            }

            // 帧命令
            array[8] = 0;
            array[9] = 0;
            array[10] = 0;
            array[11] = 0x02;

            array[12] = 0;
            array[13] = 0;
            array[14] = 0;
            array[15] = 0;

            // 命令帧头
            array[16] = 0x80; // ICF
            array[17] = 0x00; // RSV
            array[18] = 0x02; // GCT 少于8个网络节点
            array[19] = 0x00; // DNA 局域网

            array[20] = plcNode; // DA1
            array[21] = 0x00; // DA2 CPU 单元号
            array[22] = 0x00; // SNA 局域网
            array[23] = pcNode; // SA1

            array[24] = 0x00; // SA2 CPU 单元号
            array[25] = 0xFF; // SID 请求方ip

            // 指令码
            if (rw == RorW.READ)
            {
                // cmdCode--0101 读指令
                array[26] = 0x01;
                array[27] = 0x01;
            }
            else
            {
                // write---0102 写指令
                array[26] = 0x01;
                array[27] = 0x02;
            }

            // 地址
            // array[28] = (byte)mr;
            array[28] = FinsClass.GetMemoryCode(mr, mt);
            if (mr == PlcMemory.CNT)
            {
                array[29] = (byte)(ch / 256 + 128);
                array[30] = (byte)(ch % 256);
            }
            else
            {
                array[29] = (byte)(ch / 256);
                array[30] = (byte)(ch % 256);
                array[31] = (byte)offset;
            }

            array[32] = (byte)(cnt / 256);
            array[33] = (byte)(cnt % 256);

            // 返回构建的 FINS 指令字节数组
            return array;
        }

        /// <summary>
        /// 实例化PLC操作对象
        /// </summary>
        public EtherNetTcp()
        {
            client = new TcpClient();
        }

        /// <summary>
        /// 与PLC建立TCP连接
        /// </summary>
        /// <param name="ip">PLC的IP地址</param>
        /// <param name="port">端口号，一般为9600</param>
        /// <param name="timeOut">超时时间，默认3000毫秒</param>
        /// <returns>0为成功</returns>
        public short Link(string ip, int port, short timeOut = 3000)
        {
            if (timer != null)
            {
                timer.Dispose();
                timer = null;
            }

            // 连接超时
            if (!BasicClass.PingCheck(ip, timeOut)) return -1;

            client.Connect(ip, port);
            stream = client.GetStream();
            Thread.Sleep(10);

            if (SendData(FinsClass.HandShake()) != 0)
            {
                return -1;
            }

            // 开始读取返回信号
            var buffer = new byte[24];
            if (ReceiveData(buffer) != 0)
            {
                return -1;
            }

            if (buffer[15] != 0)
            {
                return -1;
            }

            pcNode = buffer[19];
            plcNode = buffer[23];

            // 启动定时器，检查连接状态
            timer = new Timer(CheckLink, null, 0, 10000);
            return 0;
        }

        /// <summary>
        /// 定时检查PLC连接状态
        /// </summary>
        private void CheckLink(Object state)
        {
            // 连接超时
            if (!BasicClass.PingCheck("150.110.60.6", 3000))
            {
                Link("150.110.60.6", 9600);
            }
        }

        /// <summary>
        /// 关闭PLC操作对象的TCP连接
        /// </summary>
        /// <returns>0为成功</returns>
        public short Close()
        {
            try
            {
                stream.Close();
                client.Close();
                return 0;
            }
            catch
            {
                return -1;
            }
        }

        /// <summary>
        /// 得到一个数据
        /// </summary>
        /// <typeparam name="T">支持：int16,int32,bool,float</typeparam>
        /// <param name="mrch">起始地址（地址：D100；位：W100.1）</param>
        /// <returns>结果值</returns>
        /// <exception cref="Exception">暂不支持此类型</exception>
        /// <exception cref="Exception">读取数据失败</exception>
        public T GetData<T>(string mrch) where T : new()
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return GetData<T>(mr, txtq);
        }

        /// <summary>
        /// 得到一个数据
        /// </summary>
        /// <typeparam name="T">支持：int16,int32,bool,float</typeparam>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">起始地址（地址：100；位：100.01）</param>
        /// <returns>结果值</returns>
        /// <exception cref="Exception">暂不支持此类型</exception>
        /// <exception cref="Exception">读取数据失败</exception>
        private T GetData<T>(PlcMemory mr, object ch) where T : new()
        {
            var t = new T();
            switch (t)
            {
                case short _:
                    if (ReadWord(mr, short.Parse(ch.ToString()), out var shortData) == 0)
                        return (T)(object)shortData;
                    break;
                case int _:
                    if (ReadInt32(mr, short.Parse(ch.ToString()), out var intData) == 0)
                        return (T)(object)intData;
                    break;
                case float _:
                    if (ReadReal(mr, short.Parse(ch.ToString()), out var floatData) == 0)
                        return (T)(object)floatData;
                    break;
                case bool _:
                    if (GetBitState(mr, ch.ToString(), out var bs) == 0)
                        return (T)(object)(bs == 1);
                    break;
                default:
                    throw new Exception("暂不支持此类型");
            }

            throw new Exception("读取数据失败");
        }

        /// <summary>
        /// 设置一个数据
        /// </summary>
        /// <typeparam name="T">支持：int16,int32,bool,float</typeparam>
        /// <param name="mrch">起始地址（地址：D100；位：W100.1）</param>
        /// <param name="inData">写入的数据</param>
        /// <returns>是否成功</returns>
        /// <exception cref="Exception">暂不支持此类型</exception>
        /// <exception cref="Exception">写入数据失败</exception>
        public void SetData<T>(string mrch, T inData) where T : new()
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            SetData(mr, txtq, inData);
        }

        /// <summary>
        /// 设置一个数据
        /// </summary>
        /// <typeparam name="T">支持：int16,int32,bool,float</typeparam>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">起始地址（地址：100；位：100.01）</param>
        /// <param name="inData">写入的数据</param>
        /// <returns>是否成功</returns>
        /// <exception cref="Exception">暂不支持此类型</exception>
        /// <exception cref="Exception">写入数据失败</exception>
        public void SetData<T>(PlcMemory mr, object ch, T inData) where T : new()
        {
            short isok;

            switch (inData)
            {
                case short dShort:
                    isok = WriteWord(mr, short.Parse(ch.ToString()), dShort);
                    break;
                case bool dBool:
                    isok = SetBitState(mr, ch.ToString(), dBool ? BitState.ON : BitState.OFF);
                    break;
                case int dInt:
                    isok = WriteInt32(mr, short.Parse(ch.ToString()), dInt);
                    break;
                case float dFloat:
                    isok = WriteReal(mr, short.Parse(ch.ToString()), dFloat);
                    break;
                default:
                    throw new Exception("暂不支持此类型");
            }

            if (isok != 0) throw new Exception("写入数据失败");
        }

        /// <summary>
        /// 得到多个数据
        /// </summary>
        /// <typeparam name="T">支持：int16,int32,bool,float</typeparam>
        /// <param name="mrch">起始地址（地址：D100；位：W100.1）</param>
        /// <param name="count">读取个数</param>
        /// <returns>结果值</returns>
        /// <exception cref="Exception">暂不支持此类型</exception>
        /// <exception cref="Exception">读取数据失败</exception>
        public T[] GetDatas<T>(string mrch, int count) where T : new()
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return GetDatas<T>(mr, txtq, count);
        }

        /// <summary>
        /// 得到多个数据
        /// </summary>
        /// <typeparam name="T">支持：int16,int32,bool,float</typeparam>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">起始地址（地址：100；位：100.01）</param>
        /// <param name="count">读取个数</param>
        /// <returns>结果值</returns>
        /// <exception cref="Exception">暂不支持此类型</exception>
        /// <exception cref="Exception">读取数据失败</exception>
        private T[] GetDatas<T>(PlcMemory mr, object ch, int count) where T : new()
        {
            var t = new T();
            switch (t)
            {
                case short _:
                    if (ReadWords(mr, short.Parse(ch.ToString()), Convert.ToInt16(count), out var shortData) == 0)
                        return (T[])(object)shortData;
                    throw new Exception("读取数据失败");

                case bool _:
                    // var ts = new T[count];
                    if (GetBitStates(mr, ch.ToString(), out var boolData, (short)count) == 0)
                        return (T[])(object)boolData;
                    throw new Exception("读取数据失败");

                case int _:
                    var intTs = new T[count];

                    var intCh = short.Parse(ch.ToString());
                    for (var i = 0; i < intTs.Length; i++)
                    {
                        if (ReadInt32(mr, Convert.ToInt16(intCh + i), out var intData) == 0)
                            intTs[i] = (T)(object)intData;
                        else
                            throw new Exception("读取数据失败");
                    }

                    return intTs;

                case float _:
                    var floatTs = new T[count];

                    var ch2 = short.Parse(ch.ToString());
                    for (var i = 0; i < floatTs.Length; i++)
                    {
                        if (ReadReal(mr, Convert.ToInt16(ch2 + i), out var floatData) == 0)
                            floatTs[i] = (T)(object)floatData;
                        else
                            throw new Exception("读取数据失败");
                    }

                    return floatTs;

                default:
                    throw new Exception("暂不支持此类型");
            }
        }

        /// <summary>
        /// 设置多个数据
        /// </summary>
        /// <typeparam name="T">支持：int16,</typeparam>
        /// <param name="mrch">起始地址（地址：D100；位：W100.1）</param>
        /// <param name="inDatas">写入的数据</param>
        /// <returns>是否成功</returns>
        /// <exception cref="Exception">暂不支持此类型</exception>
        /// <exception cref="Exception">写入数据失败</exception>
        public void SetDatas<T>(string mrch, params T[] inDatas) where T : new()
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            SetDatas(mr, txtq, inDatas);
        }

        /// <summary>
        /// 设置多个数据
        /// </summary>
        /// <typeparam name="T">支持：int16,</typeparam>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">起始地址（地址：100；位：100.01）</param>
        /// <param name="inDatas">写入的数据</param>
        /// <returns>是否成功</returns>
        /// <exception cref="Exception">暂不支持此类型</exception>
        /// <exception cref="Exception">写入数据失败</exception>
        private void SetDatas<T>(PlcMemory mr, object ch, params T[] inDatas) where T : new()
        {
            short isok;

            if (inDatas is short[] dInt)
                isok = WriteWords(mr, short.Parse(ch.ToString()), Convert.ToInt16(dInt.Length), dInt);
            else
                throw new Exception("暂不支持此类型");

            if (isok != 0) throw new Exception("写入数据失败");
        }

        /// <summary>
        /// 读值方法（多个连续值）
        /// </summary>
        /// <param name="mrch">起始地址。如：D100,W100.1</param>
        /// <param name="cnt">地址个数</param>
        /// <param name="reData">返回值</param>
        /// <returns>0为成功</returns>
        public short ReadWords(string mrch, short cnt, out short[] reData)
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return ReadWords(mr, short.Parse(txtq), cnt, out reData);
        }

        /// <summary>
        /// 读值方法（多个连续值）
        /// </summary>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">起始地址</param>
        /// <param name="cnt">地址个数</param>
        /// <param name="reData">返回值</param>
        /// <returns>0为成功</returns>
        private short ReadWords(PlcMemory mr, short ch, short cnt, out short[] reData)
        {
            reData = new short[cnt]; // 储存读取到的数据
            var num = (30 + cnt * 2); // 接收数据(Text)的长度,字节数
            var buffer = new byte[num]; // 用于接收数据的缓存区大小
            var array = FinsCmd(RorW.READ, mr, MemoryType.WORD, ch, 00, cnt);

            if (SendData(array) != 0) return -1;

            if (ReceiveData(buffer) != 0) return -1;

            // 命令返回成功，继续查询是否有错误码，然后在读取数据
            var succeed = true;
            if (buffer[11] == 3)
                succeed = ErrorCode.CheckHeadError(buffer[15]);

            // no header error
            if (!succeed) return -1;

            // endcode为fins指令的返回错误码
            if (!ErrorCode.CheckEndCode(buffer[28], buffer[29])) return -1;

            // 完全正确的返回，开始读取返回的具体数值
            for (var i = 0; i < cnt; i++)
            {
                // 返回的数据从第30字节开始储存的,
                // PLC每个字占用两个字节，且是高位在前，这和微软的默认低位在前不同
                // 因此无法直接使用，reData[i] = BitConverter.ToInt16(buffer, 30 + i * 2);
                // 先交换了高低位的位置，然后再使用BitConverter.ToInt16转换
                var temp = new[] { buffer[30 + i * 2 + 1], buffer[30 + i * 2] };
                reData[i] = BitConverter.ToInt16(temp, 0);
            }

            return 0;
        }

        /// <summary>
        /// 读单个字方法
        /// </summary>
        /// <param name="mrch">起始地址。如：D100,W100.1</param>
        /// <param name="reData">返回值</param>
        /// <returns>0为成功</returns>
        public short ReadWord(string mrch, out short reData)
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return ReadWord(mr, short.Parse(txtq), out reData);
        }

        /// <summary>
        /// 读单个字方法
        /// </summary>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">起始地址</param>
        /// <param name="reData">返回值</param>
        /// <returns>0为成功</returns>
        private short ReadWord(PlcMemory mr, short ch, out short reData)
        {
            // short[] temp;
            reData = new short();
            var re = ReadWords(mr, ch, 1, out short[] temp);
            if (re != 0)
            {
                return -1;
            }

            reData = temp[0];
            return 0;
        }

        /// <summary>
        /// 写值方法（多个连续值）
        /// </summary>
        /// <param name="mrch">起始地址。如：D100,W100.1</param>
        /// <param name="cnt">地址个数</param>
        /// <param name="inData">写入值</param>
        /// <returns>0为成功</returns>
        public short WriteWords(string mrch, short cnt, short[] inData)
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return WriteWords(mr, short.Parse(txtq), cnt, inData);
        }

        /// <summary>
        /// 写值方法（多个连续值）
        /// </summary>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">起始地址</param>
        /// <param name="cnt">地址个数</param>
        /// <param name="inData">写入值</param>
        /// <returns>0为成功</returns>
        private short WriteWords(PlcMemory mr, short ch, short cnt, short[] inData)
        {
            var buffer = new byte[30];

            // 前34字节和读指令基本一直，还需要拼接下面的输入数据数组
            var arrayhead = FinsCmd(RorW.WRITE, mr, MemoryType.WORD, ch, 00, cnt);
            var wdata = new byte[(cnt * 2)];

            // 转换写入值到wdata数组
            for (var i = 0; i < cnt; i++)
            {
                var temp = BitConverter.GetBytes(inData[i]);

                // 转换为PLC的高位在前储存方式
                wdata[i * 2] = temp[1];
                wdata[i * 2 + 1] = temp[0];
            }

            // 拼接写入数组
            var array = new byte[(cnt * 2 + 34)];
            arrayhead.CopyTo(array, 0);
            wdata.CopyTo(array, 34);
            if (SendData(array) != 0) return -1;
            if (ReceiveData(buffer) != 0) return -1;

            // 命令返回成功，继续查询是否有错误码，然后在读取数据
            var succeed = true;
            if (buffer[11] == 3)
                succeed = ErrorCode.CheckHeadError(buffer[15]);
            if (!succeed) return -1;

            // endcode为fins指令的返回错误码
            if (ErrorCode.CheckEndCode(buffer[28], buffer[29]))

                // 完全正确的返回0
                return 0;
            return -1;
        }

        /// <summary>
        /// 写单个字方法
        /// </summary>
        /// <param name="mrch">起始地址。如：D100,W100.1</param>
        /// <param name="inData">写入数据</param>
        /// <returns>0为成功</returns>
        public short WriteWord(string mrch, short inData)
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return WriteWord(mr, short.Parse(txtq), inData);
        }

        /// <summary>
        /// 写单个字方法
        /// </summary>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">地址</param>
        /// <param name="inData">写入数据</param>
        /// <returns>0为成功</returns>
        public short WriteWord(PlcMemory mr, short ch, short inData)
        {
            var temp = new[] { inData };
            var re = WriteWords(mr, ch, 1, temp);
            if (re != 0)
                return -1;
            return 0;
        }

        /// <summary>
        /// 读值方法-按位bit（单个）
        /// </summary>
        /// <param name="mrch">起始地址。如：W100.1</param>
        /// <param name="bs">返回开关状态枚举EtherNetPLC.BitState，0/1</param>
        /// <returns>0为成功</returns>
        public short GetBitState(string mrch, out short bs)
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return GetBitState(mr, txtq, out bs);
        }

        /// <summary>
        /// 读值方法-按位bit（单个）
        /// </summary>
        /// <param name="mrch">起始地址。如：W100.1</param>
        /// <param name="bs">返回开关状态枚举EtherNetPLC.BitState，0/1</param>
        /// <param name="cnt"></param>
        /// <returns>0为成功</returns>
        public short GetBitStates(string mrch, out bool[] bs, short cnt = 1)
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return GetBitStates(mr, txtq, out bs, cnt);
        }

        /// <summary>
        /// 读值方法-按位bit（单个）
        /// </summary>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">地址000.00</param>
        /// <param name="bs">返回开关状态枚举EtherNetPLC.BitState，0/1</param>
        /// <param name="cnt"></param>
        /// <returns>0为成功</returns>
        public short GetBitStates(PlcMemory mr, string ch, out bool[] bs, short cnt = 1)
        {
            bs = new bool[cnt];
            var buffer = new byte[30 + cnt]; // 用于接收数据的缓存区大小
            var cnInt = short.Parse(ch.Split('.')[0]);
            var cnBit = short.Parse(ch.Split('.')[1]);
            var array = FinsCmd(RorW.READ, mr, MemoryType.BIT, cnInt, cnBit, cnt);
            if (SendData(array) != 0) return -1;
            if (ReceiveData(buffer) != 0) return -1;

            // 命令返回成功，继续查询是否有错误码，然后在读取数据
            var succeed = true;
            if (buffer[11] == 3)
                succeed = ErrorCode.CheckHeadError(buffer[15]);
            if (!succeed) return -1;

            // endcode为fins指令的返回错误码
            if (!ErrorCode.CheckEndCode(buffer[28], buffer[29])) return -1;

            // 完全正确的返回，开始读取返回的具体数值
            for (var i = 0; i < cnt; i++) bs[i] = buffer[30 + i] == 1;

            return 0;
        }

        /// <summary>
        /// 读值方法-按位bit（单个）
        /// </summary>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">地址000.00</param>
        /// <param name="bs">返回开关状态枚举EtherNetPLC.BitState，0/1</param>
        /// <returns>0为成功</returns>
        private short GetBitState(PlcMemory mr, string ch, out short bs)
        {
            bs = new short();
            var buffer = new byte[31]; // 用于接收数据的缓存区大小
            var cnInt = short.Parse(ch.Split('.')[0]);
            var cnBit = short.Parse(ch.Split('.')[1]);
            var array = FinsCmd(RorW.READ, mr, MemoryType.BIT, cnInt, cnBit, 1);
            if (SendData(array) != 0) return -1;
            if (ReceiveData(buffer) != 0) return -1;

            // 命令返回成功，继续查询是否有错误码，然后在读取数据
            var succeed = true;
            if (buffer[11] == 3)
                succeed = ErrorCode.CheckHeadError(buffer[15]);
            if (!succeed) return -1;

            // endcode为fins指令的返回错误码
            if (!ErrorCode.CheckEndCode(buffer[28], buffer[29])) return -1;

            // 完全正确的返回，开始读取返回的具体数值
            bs = buffer[30];
            return 0;
        }

        /// <summary>
        /// 写值方法-按位bit（单个）
        /// </summary>
        /// <param name="mrch">起始地址。如：W100.1</param>
        /// <param name="bs">开关状态枚举EtherNetPLC.BitState，0/1</param>
        /// <returns>0为成功</returns>
        public short SetBitState(string mrch, BitState bs)
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return SetBitState(mr, txtq, bs);
        }

        /// <summary>
        /// 写值方法-按位bit（单个）
        /// </summary>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">地址000.00</param>
        /// <param name="bs">开关状态枚举EtherNetPLC.BitState，0/1</param>
        /// <returns>0为成功</returns>
        public short SetBitState(PlcMemory mr, string ch, BitState bs)
        {
            var buffer = new byte[30];
            var cnInt = short.Parse(ch.Split('.')[0]);
            var cnBit = short.Parse(ch.Split('.')[1]);
            var arrayhead = FinsCmd(RorW.WRITE, mr, MemoryType.BIT, cnInt, cnBit, 1);
            var array = new byte[35];
            arrayhead.CopyTo(array, 0);
            array[34] = (byte)bs;
            if (SendData(array) != 0) return -1;
            if (ReceiveData(buffer) != 0) return -1;

            // 命令返回成功，继续查询是否有错误码，然后在读取数据
            var succeed = true;
            if (buffer[11] == 3)
                succeed = ErrorCode.CheckHeadError(buffer[15]);
            if (!succeed) return -1;

            // endcode为fins指令的返回错误码
            if (ErrorCode.CheckEndCode(buffer[28], buffer[29]))

                // 完全正确的返回0
                return 0;
            return -1;
        }

        /// <summary>
        /// 读一个浮点数的方法，单精度，在PLC中占两个字
        /// </summary>
        /// <param name="mrch">起始地址，会读取两个连续的地址，因为单精度在PLC中占两个字</param>
        /// <param name="reData">返回一个float型</param>
        /// <returns>0为成功</returns>
        public short ReadReal(string mrch, out float reData)
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return ReadReal(mr, short.Parse(txtq), out reData);
        }

        /// <summary>
        /// 读一个浮点数的方法，单精度，在PLC中占两个字
        /// </summary>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">起始地址，会读取两个连续的地址，因为单精度在PLC中占两个字</param>
        /// <param name="reData">返回一个float型</param>
        /// <returns>0为成功</returns>
        private short ReadReal(PlcMemory mr, short ch, out float reData)
        {
            reData = new float();
            var num = (30 + 2 * 2); // 接收数据(Text)的长度,字节数
            var buffer = new byte[num]; // 用于接收数据的缓存区大小
            var array = FinsCmd(RorW.READ, mr, MemoryType.WORD, ch, 00, 2);
            if (SendData(array) != 0) return -1;
            if (ReceiveData(buffer) == 0) return -1;

            // 命令返回成功，继续查询是否有错误码，然后在读取数据
            var succeed = true;
            if (buffer[11] == 3)
                succeed = ErrorCode.CheckHeadError(buffer[15]);
            if (!succeed) return -1;

            // endcode为fins指令的返回错误码
            if (!ErrorCode.CheckEndCode(buffer[28], buffer[29])) return -1;

            // 完全正确的返回，开始读取返回的具体数值
            var temp = new[] { buffer[30 + 1], buffer[30], buffer[30 + 3], buffer[30 + 2] };
            reData = BitConverter.ToSingle(temp, 0);
            return 0;
        }

        /// <summary>
        /// 写一个浮点数的方法，单精度，在PLC中占两个字
        /// </summary>
        /// <param name="mrch">起始地址，会读取两个连续的地址，因为单精度在PLC中占两个字</param>
        /// <param name="reData">返回一个float型</param>
        /// <returns>0为成功</returns>
        public short WriteReal(string mrch, float reData)
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return WriteReal(mr, short.Parse(txtq), reData);
        }

        /// <summary>
        /// 写一个浮点数的方法，单精度，在PLC中占两个字
        /// </summary>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">起始地址，会读取两个连续的地址，因为单精度在PLC中占两个字</param>
        /// <param name="reData">返回一个float型</param>
        /// <returns>0为成功</returns>
        private short WriteReal(PlcMemory mr, short ch, float reData)
        {
            var temp = BitConverter.GetBytes(reData);

            var wdata = new short[] { 0, 0 };

            // TODO 待定
            if (temp.Length != 0)
                wdata[0] = BitConverter.ToInt16(temp, 0);
            if (temp.Length > 2)
                wdata[1] = BitConverter.ToInt16(temp, 2);

            var re = WriteWords(mr, ch, 2, wdata);

            return re;
        }

        /// <summary>
        /// 读一个int32的方法，在PLC中占两个字
        /// </summary>
        /// <param name="mrch">起始地址，会读取两个连续的地址，因为int32在PLC中占两个字</param>
        /// <param name="reData">返回一个int型</param>
        /// <returns>0为成功</returns>
        public short ReadInt32(string mrch, out int reData)
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return ReadInt32(mr, short.Parse(txtq), out reData);
        }

        /// <summary>
        /// 读一个int32的方法，在PLC中占两个字
        /// </summary>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">起始地址，会读取两个连续的地址，因为int32在PLC中占两个字</param>
        /// <param name="reData">返回一个int型</param>
        /// <returns>0为成功</returns>
        private short ReadInt32(PlcMemory mr, short ch, out int reData)
        {
            reData = new int();
            var num = (30 + 2 * 2); // 接收数据(Text)的长度,字节数
            var buffer = new byte[num]; // 用于接收数据的缓存区大小
            var array = FinsCmd(RorW.READ, mr, MemoryType.WORD, ch, 00, 2);
            if (SendData(array) != 0) return -1;
            if (ReceiveData(buffer) != 0) return -1;

            // 命令返回成功，继续查询是否有错误码，然后在读取数据
            var succeed = true;
            if (buffer[11] == 3)
                succeed = ErrorCode.CheckHeadError(buffer[15]);
            if (!succeed) return -1;

            // endcode为fins指令的返回错误码
            if (!ErrorCode.CheckEndCode(buffer[28], buffer[29])) return -1;

            // 完全正确的返回，开始读取返回的具体数值
            var temp = new[] { buffer[30 + 1], buffer[30], buffer[30 + 3], buffer[30 + 2] };
            reData = BitConverter.ToInt32(temp, 0);
            return 0;
        }

        /// <summary>
        /// 写一个int32的方法，在PLC中占两个字
        /// </summary>
        /// <param name="mrch">起始地址，会读取两个连续的地址，因为int32在PLC中占两个字</param>
        /// <param name="reData">返回一个int型</param>
        /// <returns>0为成功</returns>
        public short WriteInt32(string mrch, int reData)
        {
            var mr = ConvertClass.GetPlcMemory(mrch, out var txtq);
            return WriteInt32(mr, short.Parse(txtq), reData);
        }

        /// <summary>
        /// 写一个int32的方法，在PLC中占两个字
        /// </summary>
        /// <param name="mr">地址类型枚举</param>
        /// <param name="ch">起始地址，会读取两个连续的地址，因为int32在PLC中占两个字</param>
        /// <param name="reData">返回一个int型</param>
        /// <returns>0为成功</returns>
        private short WriteInt32(PlcMemory mr, short ch, int reData)
        {
            var temp = BitConverter.GetBytes(reData);

            var wdata = new short[] { 0, 0 };

            // TODO 待定
            if (temp.Length != 0)
                wdata[0] = BitConverter.ToInt16(temp, 0);
            if (temp.Length > 2)
                wdata[1] = BitConverter.ToInt16(temp, 2);

            var re = WriteWords(mr, ch, 2, wdata);

            return re;
        }
    }
}