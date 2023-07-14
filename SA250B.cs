using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Measurement
{
    public class Sa250B
    {
        // 成员变量
        // 250B接口地址
        private const string DllUrl = "C:\\Program Files (x86)\\Saunders & Associates\\250B\\W250BOLE.dll";

        // 正则规则
        private static readonly string[] Regular = { "=", "(", "),", ")" };

        // PLC驱动
        private static Omrom omrom;

        // Access数据库驱动
        private static AccessConnection accessConnection;

        // 250BPort
        private const int Port = 1;

        // 自定义公共接口
        public Task<object> Init(string port)
        {
            try
            {
                omrom = new Omrom();
                accessConnection = new AccessConnection();
                omrom.Init(port);
                accessConnection.Init();
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
            catch (Exception e)
            {
                return Task.FromResult<object>(e);
            }

            return Task.FromResult<object>("服务器启动成功");
        }

        // 关闭250B服务器
        public Task<object> CloseMeasuringProgram(object none)
        {
            try
            {
                DisconnectFromServer();
            }
            catch (Exception)
            {
                return Task.FromResult<object>("服务器关闭异常");
            }

            return Task.FromResult<object>("服务器关闭成功");
        }

        // 切换端口
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
                Measure();
            }
            catch (Exception)
            {
                return Task.FromResult<object>(false);
            }

            return Task.FromResult<object>(true);
        }

        // 测试数据
        private static string MeasurementData()
        {
            string result = null;
            try
            {
                var getResults = "";
                MeasureAndGetResultsB(1, ref getResults);
                var str = getResults.Split(Regular, StringSplitOptions.RemoveEmptyEntries);
                for (var i = 0; i < str.Length / 3; i++)
                {
                    var index = i * 3;
                    if (str[index] != "FL") continue;
                    result = str[index + 1];
                }
            }
            catch (Exception)
            {
                return null;
            }

            return result;
        }

        // 一次测试并返回数据
        public Task<object> MeasureAndReturn(object none)
        {
            var data = MeasurementData();
            return Task.FromResult(string.IsNullOrEmpty(data) ? (object)false : data);
        }

        // 一组测试并返回数据
        public Task<object> GroupTest(object none)
        {
            var dataList = new List<object>(4);
            try
            {
                for (var i = 0; i < 4; i++)
                {
                    var data = MeasurementData();
                    if (string.IsNullOrEmpty(data)) return Task.FromResult<object>(false);
                    dataList.Add(data);
                }
            }
            catch (Exception)
            {
                return Task.FromResult<object>(false);
            }

            return Task.FromResult<object>(dataList);
        }

        // 设置串口
        public Task<object> SetSerialPort(string post)
        {
            try
            {
                omrom.SetPort(post);
            }
            catch (Exception)
            {
                return Task.FromResult<object>(false);
            }

            return Task.FromResult<object>(true);
        }

        // 获取串口
        public Task<object> GetSerialPort(object none)
        {
            try
            {
                return Task.FromResult<object>(omrom.GetPort());
            }
            catch (Exception)
            {
                return Task.FromResult<object>(null);
            }
        }

        // 获取串口列表
        public Task<object> GetSerialPortList(object none)
        {
            try
            {
                return Task.FromResult<object>(omrom.GetPortList());
            }
            catch (Exception)
            {
                return Task.FromResult<object>(false);
            }
        }

        // 查询标品状态
        public Task<object> GetStandardProductData(dynamic data)
        {
            return Task.FromResult<object>(accessConnection.GetStandard(data));
        }

        // 丝杆动作
        public Task<object> ScrewAction(int i)
        {
            if (i == 0)
                omrom.UpperScrewAction();
            else
                omrom.LowerScrewAction();

            return Task.FromResult<object>(true);
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
            catch (Exception e)
            {
                return Task.FromResult<object>(false);
            }
        }

        // 校准分压
        private static bool CalibrateDivider(int steps, int portIndex, string fixture)
        {
            var success = 0;
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

            return success != 0;
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
        private static extern int Measure(int nMeasure250B1, int nMeasure250B2, int nMeasure250B3, int nMeasure250B4, ref int pn250BsTriggered);

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
        private static extern int MeasureAndGetResultsB(int nMeasure250B1AndGetResults, ref string pbstr250B1Results);

        // 测量并返回数据-指定250B
        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int MeasureAndGetResultsB(int nMeasure250B1AndGetResults, ref string pbstr250B1Results, int nMeasure250B2AndGetResults,
            ref string pbstr250B2Results, int nMeasure250B3AndGetResults, ref string pbstr250B3Results, int nMeasure250B4AndGetResults,
            ref string pbstr250B4Results);

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
        private static extern int DefineDLDSweepB(int nDLDSweepID, int nSweepIndex, double fBeginPower, string bstrBeginPowerUnit, double fEndPower,
            string bstrEndPowerUnit, int nSteps, double fDelay, int nAverages);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int DefineDLDSweepC(int nDLDSweepID, int nSweepIndex, double fBeginPower, string strBeginPowerUnit, double fEndPower,
            string strEndPowerUnit, int nSteps, double fDelay, int nAverages);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int DisableGraphUpdate(double fSecondsToDisable);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int EnableOEMKey(int nOEMNumber);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int FileOpenB(string bstrQCCFilePath);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int FileOpenC(string strQCCFilePath);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int Get250BRevision(ref double pfRevision);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int GetBinNumber(int nGet250B1Bin, ref int pn250B1Bin, int nGet250B2Bin, ref int pn250B2Bin, int nGet250B3Bin, ref int pn250B3Bin,
            int nGet250B4Bin, ref int pn250B4Bin);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int GetDLDResults(ref int pnDLDResultsID, int nGet250B1DLDResults, int nGet250B2DLDResults, int nGet250B3DLDResults,
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
        private static extern int GetResultsB(int nGet250B1Results, ref string pbstr250B1Results, int nGet250B2Results, ref string pbstr250B2Results,
            int nGet250B3Results, ref string pbstr250B3Results, int nGet250B4Results, ref string pbstr250B4Results);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int GetResultsC(int nGet250B1Results, ref string pstr250B1Results, int nGet250B2Results, ref string pstr250B2Results,
            int nGet250B3Results, ref string pstr250B3Results, int nGet250B4Results, ref string pstr250B4Results);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int GetSpurResults(ref int pnSpurResultsID, int nGet250B1SpurResults, int nGet250B2SpurResults, int nGet250B3SpurResults,
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
        public static extern int MeasureAndGetBinNumber(int nMeasureAndGet250B1Bin, ref int pn250B1Bin, int nMeasureAndGet250B2Bin, ref int pn250B2Bin,
            int nMeasureAndGet250B3Bin, ref int pn250B3Bin, int nMeasureAndGet250B4Bin, ref int pn250B4Bin);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int MeasureAndGetResultsC(int nMeasure250B1AndGetResults, ref string pstr250B1Results, int nMeasure250B2AndGetResults,
            ref string pstr250B2Results, int nMeasure250B3AndGetResults, ref string pstr250B3Results, int nMeasure250B4AndGetResults,
            ref string pstr250B4Results);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int MeasureRaw(int nAverages, int nMeasure250B1, double f250B1Frequency, double f250B1dBm, ref double pf250B1dB,
            ref double pf250B1Phase, int nMeasure250B2, double f250B2Frequency, double f250B2dBm, ref double pf250B2dB, ref double pf250B2Phase,
            int nMeasure250B3, double f250B3Frequency, double f250B3dBm, ref double pf250B3dB, ref double pf250B3Phase, int nMeasure250B4,
            double f250B4Frequency, double f250B4dBm, ref double pf250B4dB, ref double pf250B4Phase);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int MeasureRX(int nAverages, int nMeasure250B1, double f250B1Frequency, double f250B1dBm, ref double pf250B1R, ref double pf250B1X,
            int nMeasure250B2, double f250B2Frequency, double f250B2dBm, ref double pf250B2R, ref double pf250B2X, int nMeasure250B3, double f250B3Frequency,
            double f250B3dBm, ref double pf250B3R, ref double pf250B3X, int nMeasure250B4, double f250B4Frequency, double f250B4dBm, ref double pf250B4R,
            ref double pf250B4X);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int RetrieveDLDResults(int nDLDResultsID, int nStepIndex, ref double pf250B1dBm, ref double pf250B1Power,
            ref double pf250B1Resistance, ref double pf250B1Frequency, ref double pf250B2dBm, ref double pf250B2Power, ref double pf250B2Resistance,
            ref double pf250B2Frequency, ref double pf250B3dBm, ref double pf250B3Power, ref double pf250B3Resistance, ref double pf250B3Frequency,
            ref double pf250B4dBm, ref double pf250B4Power, ref double pf250B4Resistance, ref double pf250B4Frequency, ref int pnSuccess);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int RetrieveDLDSteps(int nDLDResultsID, int nRetrieve250B1Steps, ref int pn250B1Steps, int nRetrieve250B2Steps,
            ref int pn250B2Steps, int nRetrieve250B3Steps, ref int pn250B3Steps, int nRetrieve250B4Steps, ref int pn250B4Steps);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int RetrieveDLDSweepResistance(int nDLDSweepID, ref double pfResistance, ref int pnSuccess);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int RetrieveDLDSweepB(int nDLDSweepID, int nSweepIndex, ref double pfBeginPower, ref string pbstrBeginPowerUnit,
            ref double pfEndPower, ref string pbstrEndPowerUnit, ref int pnSteps, ref double pfDelay, ref int pnAverages, ref int pnSuccess);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int RetrieveDLDSweepC(int nDLDSweepID, int nSweepIndex, ref double pfBeginPower, ref string pstrBeginPowerUnit,
            ref double pfEndPower, ref string pstrEndPowerUnit, ref int pnSteps, ref double pfDelay, ref int pnAverages, ref int pnSuccess);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int RetrieveSpurSteps(int nSpurResultsID, int n250BIndex, int nSweepIndex, ref int pnSteps);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int RetrieveSpurResults(int nSpurResultsID, int n250BIndex, int nSweepIndex, int nStepIndex, ref double pfFrequency,
            ref double pfAmplitude, ref double pfR1, ref double pfX1, ref double pfZAmplitude, ref double pfZPhase, ref int pnSuccess);

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
        private static extern int SetSerialNumbersB(int nSet250B1SerialNumber, string bstr250B1SerialNumber, int nSet250B2SerialNumber,
            string bstr250B2SerialNumber, int nSet250B3SerialNumber, string bstr250B3SerialNumber, int nSet250B4SerialNumber, string bstr250B4SerialNumber,
            ref int pnSerialNumbersSet);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SetSerialNumbersC(int nSet250B1SerialNumber, string str250B1SerialNumber, int nSet250B2SerialNumber,
            string str250B2SerialNumber, int nSet250B3SerialNumber, string str250B3SerialNumber, int nSet250B4SerialNumber, string str250B4SerialNumber,
            ref int pnSerialNumbersSet);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SimplifiedOLEInterfaceB(string bstrCommand, string bstrParameters, ref string pbstrResult);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int SimplifiedOLEInterfaceC(string strCommand, string strParameters, ref string pstrResult);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        private static extern int VerifyCalibrationC(int nPromptForLoadResistor, ref string pstrResult);

        [DllImport(DllUrl, CharSet = CharSet.Unicode)]
        public static extern int WaitForMeasurementComplete();
    }
}