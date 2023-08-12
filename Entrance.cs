using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Measurement.Access;
using Measurement.Automation;
using Measurement.FinsTcp;

namespace Measurement
{
    public class Entrance
    {
        // 成员变量
        // 250B接口地址
        private const string DllUrl = "C:\\Program Files (x86)\\Saunders & Associates\\250B\\W250BOLE.dll";

        // 正则规则
        private static readonly string[] Regular = { "=", "(", "),", ")" };

        // Access数据库驱动
        private static AccessConnection accessConnection;

        // 250BPort
        private const int Port = 1;

        // PLC驱动
        private static EtherNetTcp ent;

        // 监听线程
        private static Thread listenThread;

        // UI控制
        private static AutoUi autoUi;

        // 自定义公共接口
        // 初始化
        public Task<object> Init(string port)
        {
            try
            {
                accessConnection = new AccessConnection();
                autoUi = new AutoUi();
                accessConnection.Init();
                ent = new EtherNetTcp();
                ent.Link("150.110.60.6", 9600);
            }
            catch (Exception e)
            {
                return Task.FromResult<object>(e);
            }

            return Task.FromResult<object>("初始化成功");
        }

        // 打开250B服务器
        public Task<object> OpenMeasuringProgram(object none)
        {
            try
            {
                ConnectToServer();
            }
            catch
            {
                return Task.FromResult<object>(false);
            }

            return Task.FromResult<object>(true);
        }

        // 关闭250B服务器
        public Task<object> CloseMeasuringProgram(object none)
        {
            try
            {
                DisconnectFromServer();
                ent.Close();
            }
            catch (Exception)
            {
                return Task.FromResult<object>("服务器关闭异常");
            }

            return Task.FromResult<object>("服务器关闭成功");
        }

        // 切换量测端口
        public Task<object> SwitchPort(int port)
        {
            try
            {
                SelectPort(1, port);
            }
            catch (Exception)
            {
                return Task.FromResult<object>(false);
            }

            return Task.FromResult<object>(true);
        }

        // 一次测试
        public Task<object> SingleTest(object none)
        {
            try
            {
                return Task.FromResult<object>(MeasurementDataAll());
            }
            catch (Exception)
            {
                return Task.FromResult<object>(false);
            }
        }

        // 测试数据
        private static string MeasurementData()
        {
            string results = null;
            try
            {
                var result1 = "";
                var result2 = "";
                var result3 = "";
                var result4 = "";
                MeasureAndGetResultsB(1, ref result1, 0, ref result2, 0, ref result3, 0, ref result4);
                var str = result1.Split(Regular, StringSplitOptions.RemoveEmptyEntries);
                for (var i = 0; i < str.Length / 3; i++)
                {
                    var index = i * 3;
                    if (str[index] != "FL") continue;
                    results = str[index + 1];
                }
            }
            catch (Exception)
            {
                return null;
            }

            return results;
        }

        // 一次测试并返回数据
        public Task<object> MeasureAndReturn(object none)
        {
            var data = MeasurementData();
            return Task.FromResult(string.IsNullOrEmpty(data) ? (object)null : data);
        }

        // 一组测试并返回数据
        public Task<object> TestOneGroup(object none)
        {
            var dataList = new List<string>(4);
            try
            {
                for (var i = 0; i < 4; i++)
                {
                    SelectPort(1, i);
                    var data = MeasurementData();
                    if (string.IsNullOrEmpty(data)) return Task.FromResult<object>(null);
                    dataList.Add(data);
                }
            }
            catch (Exception)
            {
                return Task.FromResult<object>(null);
            }

            return Task.FromResult<object>(dataList);
        }

        // 一次测试并返回全部数据
        private List<object> MeasurementDataAll()
        {
            var results = new List<object>();
            try
            {
                var result1 = "";
                var result2 = "";
                var result3 = "";
                var result4 = "";
                MeasureAndGetResultsB(1, ref result1, 0, ref result2, 0, ref result3, 0, ref result4);
                var str = result1.Split(Regular, StringSplitOptions.RemoveEmptyEntries);
                for (var i = 0; i < str.Length / 3; i++)
                {
                    var index = i * 3;
                    var ret = new[] { str[index], str[index + 1], str[index + 2] };
                    results.Add(ret);
                }
            }
            catch (Exception)
            {
                return null;
            }

            return results;
        }

        // 一组测试并返回全部数据
        public Task<object> TestOneGroupData(object none)
        {
            var dataList = new List<object>(4);
            try
            {
                for (var i = 0; i < 4; i++)
                {
                    SelectPort(1, i);
                    var data = MeasurementDataAll();

                    if (data == null) return Task.FromResult<object>(null);
                    dataList.Add(data);
                }
            }
            catch (Exception)
            {
                return Task.FromResult<object>(null);
            }

            return Task.FromResult<object>(dataList);
        }

        // 查询标品状态
        public Task<object> GetStandardProductData(dynamic data)
        {
            return Task.FromResult<object>(accessConnection.GetStandard(data));
        }

        // 校机
        public Task<object> Proofreading(dynamic data)
        {
            try
            {
                var steps = (int)data.steps;
                var portIndex = (int)data.portIndex;
                var fixture = (string)data.fixture;
                return Task.FromResult<object>(CalibrateDivider(steps, portIndex, fixture));
            }
            catch (Exception)
            {
                return Task.FromResult<object>(false);
            }
        }

        // 刷新应用程序
        public Task<object> RefreshApplication(string number)
        {
            try
            {
                return Task.FromResult<object>(autoUi.VerificationState(number));
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }

        // 校准分压
        private static bool CalibrateDivider(int steps, int portIndex, string fixture)
        {
            try
            {
                StartTimeoutListener();
                var success = 0;
                Thread.Sleep(50);
                switch (steps)
                {
                    case 0:
                        CalibrateShortB(Port, portIndex, fixture, ref success);
                        break;
                    case 1:
                        CalibrateLoadB(Port, portIndex, fixture, ref success);
                        break;
                    case 2:
                        CalibrateOpenB(Port, portIndex, fixture, ref success);
                        break;
                }

                StopTimeoutListener();
                return success != 0;
            }
            catch
            {
                StopTimeoutListener();
                return false;
            }
        }

        // 保存校准数据
        public Task<object> SaveCalibrationData(object none)
        {
            try
            {
                for (var i = 0; i < 4; i++)
                {
                    var start = 0;
                    SaveCalibration(1, i, ref start);
                    if (start == 0)
                    {
                        return Task.FromResult<object>(false);
                    }
                }

                return Task.FromResult<object>(true);
            }
            catch (Exception)
            {
                return Task.FromResult<object>(false);
            }
        }

        // 写入标品补偿值
        public Task<object> WriteStandardProduct(string data)
        {
            try
            {
                ClearAllMeasurements();
                var b = accessConnection.AllChange(data);
                ClearAllMeasurements();
                return Task.FromResult<object>(b);
            }
            catch (Exception e)
            {
                return Task.FromResult<object>(e);
            }
        }

        // 读取测试上下限
        public Task<object> GetTestRestrict(object none)
        {
            return Task.FromResult<object>(accessConnection.GetTestSet());
        }

        // 改变文件
        public Task<object> ChangeFile(string path)
        {
            try
            {
                FileOpenB(path);
                return Task.FromResult<object>(true);
            }
            catch (Exception)
            {
                return Task.FromResult<object>(false);
            }
        }

        // 设置运行模式
        private bool SetRunMode(int mode)
        {
            try
            {
                switch (mode)
                {
                    // 手动模式
                    case 0:
                        ent.WriteWord(PlcMemory.CIO, 4604, 1);
                        break;

                    // 自动模式
                    case 1:
                        ent.WriteWord(PlcMemory.CIO, 4604, 2);
                        break;

                    // 校机模式
                    case 2:
                        ent.WriteWord(PlcMemory.CIO, 4604, 4);
                        break;

                    // 对机模式
                    case 3:
                        ent.WriteWord(PlcMemory.CIO, 4604, 8);
                        break;

                    // 试调模式
                    case 4:
                        ent.WriteWord(PlcMemory.CIO, 4604, 16);
                        break;

                    // 校对机模式
                    case 5:
                        ent.WriteWord(PlcMemory.CIO, 4604, 12);
                        break;

                    // 清空模式
                    case 6:
                        ent.WriteWord(PlcMemory.CIO, 4604, 0);
                        break;
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // 开启自动模式
        public Task<object> OpenAutoMode(object none)
        {
            return SetRunMode(1) ? Task.FromResult<object>(true) : Task.FromResult<object>(false);
        }

        // 校对机模式
        public Task<object> OpenProofreadingMode(int index)
        {
            try
            {
                switch (index)
                {
                    case 0:
                        SetRunMode(5);
                        break;
                    case 1:
                        SetRunMode(2);
                        break;
                    case 2:
                        SetRunMode(3);
                        break;
                }

                SetRunMode(6);
                return Task.FromResult<object>(true);
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }

        // 测试头位置
        public Task<object> TestHeadPosition(int i)
        {
            try
            {
                switch (i)
                {
                    // 测试头间距移动
                    case 0:
                        ent.WriteWord(PlcMemory.CIO, 4604, 65);
                        break;

                    // 测试头气缸下降
                    case 1:
                        ent.WriteWord(PlcMemory.CIO, 4604, 129);
                        break;

                    // 测试头气缸上升
                    case 2:
                        ent.WriteWord(PlcMemory.CIO, 4604, 257);
                        break;

                    // 回原点
                    case 3:
                        ent.WriteWord(PlcMemory.CIO, 4604, 513);
                        break;
                }

                // 状态刷新
                return Task.FromResult<object>(SetRunMode(6));
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }


        // 手动位置控制
        public Task<object> ManualPosition(int i)
        {
            try
            {
                SetRunMode(0);
                switch (i)
                {
                    // 准备位置0
                    case 0:
                        ent.WriteWord(PlcMemory.CIO, 4603, 1);
                        break;

                    // 准备位置1
                    case 1:
                        ent.WriteWord(PlcMemory.CIO, 4603, 2);
                        break;

                    // 准备位置2
                    case 2:
                        ent.WriteWord(PlcMemory.CIO, 4603, 4);
                        break;

                    // 准备位置3
                    case 3:
                        ent.WriteWord(PlcMemory.CIO, 4603, 8);
                        break;

                    // 准备位置4
                    case 4:
                        ent.WriteWord(PlcMemory.CIO, 4603, 16);
                        break;

                    // 准备位置5
                    case 5:
                        ent.WriteWord(PlcMemory.CIO, 4603, 32);
                        break;

                    // 准备位置6
                    case 6:
                        ent.WriteWord(PlcMemory.CIO, 4603, 64);
                        break;

                    // 校机位置
                    case 7:
                        ent.WriteWord(PlcMemory.CIO, 4603, 256);
                        break;

                    // 对机位置
                    case 8:
                        ent.WriteWord(PlcMemory.CIO, 4603, 512);
                        break;
                }

                ent.WriteWord(PlcMemory.CIO, 4603, 0);

                return Task.FromResult<object>(true);
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }

        // 写入位置上下限
        public Task<object> WritePositionLimit(dynamic obj)
        {
            try
            {
                var index = (bool)obj.index;
                var count = (int)obj.count;
                if (index)
                {
                    ent.SetData(PlcMemory.DM, 404, count);
                    return Task.FromResult<object>(true);
                }

                ent.SetData(PlcMemory.DM, 406, count);
                return Task.FromResult<object>(true);
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }

        // 读取量测开始信号
        public Task<object> ReadMeasureStart(object none)
        {
            try
            {
                ent.GetBitStates(PlcMemory.CIO, "4601.05", out var state);
                Thread.Sleep(80);
                return Task.FromResult<object>(state[0]);
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }

        // 输出量测完毕信号
        public Task<object> WriteMeasureEnd(object none)
        {
            try
            {
                ent.SetBitState(PlcMemory.CIO, "4601.05", BitState.OFF);
                ent.SetBitState(PlcMemory.CIO, "4601.06", BitState.ON);
                Thread.Sleep(20);
                ent.SetBitState(PlcMemory.CIO, "4601.06", BitState.OFF);
                return Task.FromResult<object>(true);
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }

        // 输出开路完成信号
        public Task<object> WriteOpenEnd(object none)
        {
            try
            {
                ent.SetBitState(PlcMemory.CIO, "4607.14 ", BitState.ON);
                Thread.Sleep(20);
                ent.SetBitState(PlcMemory.CIO, "4607.14 ", BitState.OFF);
                return Task.FromResult<object>(true);
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }

        // 关闭测试
        public Task<object> CloseTest(object none)
        {
            try
            {
                ent.WriteWord(PlcMemory.CIO, 4604, 513);
                var state = true;
                while (state)
                {
                    ent.GetBitStates(PlcMemory.CIO, "4601.10", out var states);
                    var b = states[0];
                    if (b)
                    {
                        state = false;
                    }

                    Thread.Sleep(50);
                }

                ent.WriteWord(PlcMemory.CIO, 4604, 0);
                return Task.FromResult<object>(true);
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }

        // 错误终止
        public Task<object> ErrorStop(object none)
        {
            try
            {
                ent.SetBitState(PlcMemory.CIO, "4612.00", BitState.ON);
                Thread.Sleep(100);
                ent.SetBitState(PlcMemory.CIO, "4612.00", BitState.OFF);
                return Task.FromResult<object>(true);
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }

        // 开启超时监听
        private static void StartTimeoutListener()
        {
            listenThread = new Thread(autoUi.WaitMainThreadTimeout);
            listenThread.Start();
        }

        // 关闭超时监听
        private static void StopTimeoutListener()
        {
            autoUi.StopToken();
            listenThread = null;
        }

        // 保存文件
        public Task<object> SaveFile(object none)
        {
            return Task.FromResult<object>(autoUi.SaveData());
        }

        // 当前行
        public Task<object> CurrentPosition(object none)
        {
            try
            {
                ent.ReadWords("I4605.01", 6, out var data);
                for (int i = 0; i < data.Length; i++)
                {
                    if (data[i] != 0)
                    {
                        return Task.FromResult<object>(i * 4);
                    }
                }

                return Task.FromResult<object>(null);
            }
            catch
            {
                return Task.FromResult<object>(null);
            }
        }

        // 当前列
        public Task<object> CurrentColumn(object none)
        {
            try
            {
                return Task.FromResult<object>(CalculateColumnMap(ent.GetData<int>("D2202")));
            }
            catch
            {
                return Task.FromResult<object>(null);
            }
        }

        // 计算对应列映射
        private int CalculateColumnMap(int input)
        {
            try
            {
                return input * 2 + 5;
            }
            catch
            {
                return 0;
            }
        }

        // 清除
        public Task<object> Clear(object none)
        {
            try
            {
                ClearAllMeasurements();
                return Task.FromResult<object>(true);
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }

        // 开始报警
        public Task<object> StartAlarm(object none)
        {
            try
            {
                ent.SetBitState(PlcMemory.CIO, "4614.00", BitState.ON);
                return Task.FromResult<object>(true);
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }

        // 停止报警
        public Task<object> StopAlarm(object none)
        {
            try
            {
                ent.SetBitState(PlcMemory.CIO, "4614.00", BitState.OFF);
                return Task.FromResult<object>(true);
            }
            catch
            {
                return Task.FromResult<object>(false);
            }
        }

        // 查询盘号
        public Task<object> QueryDisk(object none)
        {
            try
            {
                ent.ReadWords("H3.00", 5, out var data1);
                for (var i = 0; i < data1.Length; i++)
                {
                    if (data1[i] != 0)
                    {
                        return Task.FromResult<object>(i + 1);
                    }
                }

                ent.ReadWords("H4.00", 5, out var data2);
                for (var j = 0; j < data2.Length; j++)
                {
                    if (data2[j] != 0)
                    {
                        return Task.FromResult<object>(j + 6);
                    }
                }

                return Task.FromResult<object>(null);
            }
            catch
            {
                return Task.FromResult<object>(null);
            }
        }


        // 250B公共接口

        // 启动250B服务器
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int ConnectToServer();

        // 停止250B服务器
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int DisconnectFromServer();

        // 端口切换
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SelectPort(int nIsPortA, int nPortIndex);

        // 单次测试
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int Measure();

        // 单次测试-指定250B
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int Measure(int nMeasure250B1, int nMeasure250B2, int nMeasure250B3, int nMeasure250B4,
            ref int pn250BsTriggered);

        // 校准短路界面
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int CalibrateShortB(int nIsPortA, int nPortIndex, string bstrFixture, ref int pnSuccess);

        // 校准负载界面
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int CalibrateLoadB(int nIsPortA, int nPortIndex, string bstrFixture, ref int pnSuccess);

        // 校准开路界面 
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int CalibrateOpenB(int nIsPortA, int nPortIndex, string bstrFixture, ref int pnSuccess);

        // 保存校准
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SaveCalibration(int nIsPortA, int nPortIndex, ref int pnSuccess);

        // 取消校准
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int CancelCalibration(int nIsPortA, int nPortIndex, ref int pnSuccess);

        // 清除测量
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int ClearAllMeasurements();

        // 验证校准
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int VerifyCalibrationB(int nPromptForLoadResistor, ref string pbstrResult);

        // 测量并返回数据
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int MeasureAndGetResultsB(int nMeasure250B1AndGetResults,
            [MarshalAs(UnmanagedType.BStr)] ref string pbstr250B1Results);

        // 测量并返回数据-指定250B
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int MeasureAndGetResultsB(int nMeasure250B1AndGetResults,
            [MarshalAs(UnmanagedType.BStr)] ref string pbstr250B1Results, int nMeasure250B2AndGetResults,
            [MarshalAs(UnmanagedType.BStr)] ref string pbstr250B2Results, int nMeasure250B3AndGetResults,
            [MarshalAs(UnmanagedType.BStr)] ref string pbstr250B3Results, int nMeasure250B4AndGetResults,
            [MarshalAs(UnmanagedType.BStr)] ref string pbstr250B4Results);

        // 打开或改变QCC文件
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int FileOpenB(string bstrQccFilePath);

        // .............................

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int CalibrateLoadC(int nIsPortA, int nPortIndex, string strFixture, ref int pnSuccess);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int CalibrateOpenC(int nIsPortA, int nPortIndex, string strFixture, ref int pnSuccess);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int CalibrateShortC(int nIsPortA, int nPortIndex, string strFixture, ref int pnSuccess);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int ClearReferenceStandardError();

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int CompactDatabase();

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int CreateDLDSweep(ref int pnDLDSweepID, int nSweeps, double fResistance);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int DefineDLDSweepB(int nDLDSweepID, int nSweepIndex, double fBeginPower,
            string bstrBeginPowerUnit, double fEndPower,
            string bstrEndPowerUnit, int nSteps, double fDelay, int nAverages);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int DefineDLDSweepC(int nDLDSweepID, int nSweepIndex, double fBeginPower,
            string strBeginPowerUnit, double fEndPower,
            string strEndPowerUnit, int nSteps, double fDelay, int nAverages);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int DisableGraphUpdate(double fSecondsToDisable);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int EnableOEMKey(int nOEMNumber);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int FileOpenC(string strQCCFilePath);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int Get250BRevision(ref double pfRevision);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int GetBinNumber(int nGet250B1Bin, ref int pn250B1Bin, int nGet250B2Bin,
            ref int pn250B2Bin, int nGet250B3Bin, ref int pn250B3Bin,
            int nGet250B4Bin, ref int pn250B4Bin);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int GetDLDResults(ref int pnDLDResultsID, int nGet250B1DLDResults, int nGet250B2DLDResults,
            int nGet250B3DLDResults,
            int nGet250B4DLDResults);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int GetDLDSweep(ref int pnDLDSweepID, ref int pnSweeps);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int GetDLLRevision(ref double pfRevision);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int GetNamePlateDataB(ref string pbstrNamePlateData);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int GetNamePlateDataC(ref string pstrNamePlateData);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int GetReferenceStandardStatusB(ref string pbstrReferenceStandardStatus);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int GetReferenceStandardStatusC(ref string pstrReferenceStandardStatus);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int GetResultsB(int nGet250B1Results, ref string pbstr250B1Results, int nGet250B2Results,
            ref string pbstr250B2Results,
            int nGet250B3Results, ref string pbstr250B3Results, int nGet250B4Results, ref string pbstr250B4Results);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int GetResultsC(int nGet250B1Results, ref string pstr250B1Results, int nGet250B2Results,
            ref string pstr250B2Results,
            int nGet250B3Results, ref string pstr250B3Results, int nGet250B4Results, ref string pstr250B4Results);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int GetSpurResults(ref int pnSpurResultsID, int nGet250B1SpurResults,
            int nGet250B2SpurResults, int nGet250B3SpurResults,
            int nGet250B4SpurResults);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int GetTestTypesB(ref string pbstrTestTypes);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int GetTestTypesC(ref string pstrTestTypes);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int IsConfigurationDifferent(ref int pnConfigurationisDifferent);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int IsMeasuring(ref int pnIsMeasuring);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int IsReferenceStandardOK(ref int pnReferenceStandardIsOK);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int LockUserMode(int nLock);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int MeasureAndGetBinNumber(int nMeasureAndGet250B1Bin, ref int pn250B1Bin,
            int nMeasureAndGet250B2Bin, ref int pn250B2Bin,
            int nMeasureAndGet250B3Bin, ref int pn250B3Bin, int nMeasureAndGet250B4Bin, ref int pn250B4Bin);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int MeasureAndGetResultsC(int nMeasure250B1AndGetResults, ref string pstr250B1Results,
            int nMeasure250B2AndGetResults,
            ref string pstr250B2Results, int nMeasure250B3AndGetResults, ref string pstr250B3Results,
            int nMeasure250B4AndGetResults,
            ref string pstr250B4Results);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int MeasureRaw(int nAverages, int nMeasure250B1, double f250B1Frequency, double f250B1dBm,
            ref double pf250B1dB,
            ref double pf250B1Phase, int nMeasure250B2, double f250B2Frequency, double f250B2dBm, ref double pf250B2dB,
            ref double pf250B2Phase,
            int nMeasure250B3, double f250B3Frequency, double f250B3dBm, ref double pf250B3dB, ref double pf250B3Phase,
            int nMeasure250B4,
            double f250B4Frequency, double f250B4dBm, ref double pf250B4dB, ref double pf250B4Phase);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int MeasureRX(int nAverages, int nMeasure250B1, double f250B1Frequency, double f250B1dBm,
            ref double pf250B1R, ref double pf250B1X,
            int nMeasure250B2, double f250B2Frequency, double f250B2dBm, ref double pf250B2R, ref double pf250B2X,
            int nMeasure250B3, double f250B3Frequency,
            double f250B3dBm, ref double pf250B3R, ref double pf250B3X, int nMeasure250B4, double f250B4Frequency,
            double f250B4dBm, ref double pf250B4R,
            ref double pf250B4X);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int RetrieveDLDResults(int nDLDResultsID, int nStepIndex, ref double pf250B1dBm,
            ref double pf250B1Power,
            ref double pf250B1Resistance, ref double pf250B1Frequency, ref double pf250B2dBm, ref double pf250B2Power,
            ref double pf250B2Resistance,
            ref double pf250B2Frequency, ref double pf250B3dBm, ref double pf250B3Power, ref double pf250B3Resistance,
            ref double pf250B3Frequency,
            ref double pf250B4dBm, ref double pf250B4Power, ref double pf250B4Resistance, ref double pf250B4Frequency,
            ref int pnSuccess);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int RetrieveDLDSteps(int nDLDResultsID, int nRetrieve250B1Steps, ref int pn250B1Steps,
            int nRetrieve250B2Steps,
            ref int pn250B2Steps, int nRetrieve250B3Steps, ref int pn250B3Steps, int nRetrieve250B4Steps,
            ref int pn250B4Steps);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int
            RetrieveDLDSweepResistance(int nDLDSweepID, ref double pfResistance, ref int pnSuccess);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int RetrieveDLDSweepB(int nDLDSweepID, int nSweepIndex, ref double pfBeginPower,
            ref string pbstrBeginPowerUnit,
            ref double pfEndPower, ref string pbstrEndPowerUnit, ref int pnSteps, ref double pfDelay,
            ref int pnAverages, ref int pnSuccess);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int RetrieveDLDSweepC(int nDLDSweepID, int nSweepIndex, ref double pfBeginPower,
            ref string pstrBeginPowerUnit,
            ref double pfEndPower, ref string pstrEndPowerUnit, ref int pnSteps, ref double pfDelay, ref int pnAverages,
            ref int pnSuccess);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int RetrieveSpurSteps(int nSpurResultsID, int n250BIndex, int nSweepIndex,
            ref int pnSteps);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int RetrieveSpurResults(int nSpurResultsID, int n250BIndex, int nSweepIndex,
            int nStepIndex, ref double pfFrequency,
            ref double pfAmplitude, ref double pfR1, ref double pfX1, ref double pfZAmplitude, ref double pfZPhase,
            ref int pnSuccess);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int SetCalculatedFLMode();

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int SetDLDSweep(int nDLDSweepID, ref int pnSweepsSet);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int SetMeasuredFLMode();

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SetNamePlateDataB(string bstrNamePlateData, ref int pnParametersSet);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SetNamePlateDataC(string strNamePlateData, ref int pnParametersSet);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int SetPhysicalFLMode(double fPhysicalLoad);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int SetReferenceCL(double fReferenceCL);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int SetReferenceFR(double fReferenceFR);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SetReferencePowerB(double fPower, string bstrUnits, double fResistance);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SetReferencePowerC(double fPower, string strUnits, double fResistance);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SetSerialNumbersB(int nSet250B1SerialNumber, string bstr250B1SerialNumber,
            int nSet250B2SerialNumber,
            string bstr250B2SerialNumber, int nSet250B3SerialNumber, string bstr250B3SerialNumber,
            int nSet250B4SerialNumber, string bstr250B4SerialNumber,
            ref int pnSerialNumbersSet);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SetSerialNumbersC(int nSet250B1SerialNumber, string str250B1SerialNumber,
            int nSet250B2SerialNumber,
            string str250B2SerialNumber, int nSet250B3SerialNumber, string str250B3SerialNumber,
            int nSet250B4SerialNumber, string str250B4SerialNumber,
            ref int pnSerialNumbersSet);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SimplifiedOLEInterfaceB(string bstrCommand, string bstrParameters,
            ref string pbstrResult);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SimplifiedOLEInterfaceC(string strCommand, string strParameters,
            ref string pstrResult);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int VerifyCalibrationC(int nPromptForLoadResistor, ref string pstrResult);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int WaitForMeasurementComplete();
    }
}