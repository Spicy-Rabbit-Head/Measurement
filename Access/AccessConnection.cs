using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using Measurement.Entity;

namespace Measurement.Access
{
    internal class AccessConnection
    {
        // 标志地址
        private const string Mark = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";

        // 250B补正地址
        private const string Compensate = "C:\\ProgramData\\Saunders & Associates\\250B\\Data\\Server-1.mdb";

        // 标品地址
        private const string StandardArticle = ";jet oledb:Database Password=";

        // Access 连接实例
        private OleDbConnection ole;

        // 改变单个补正值
        public bool Change(double num)
        {
            var sql = $"update Tests set ReferenceValue='{num.ToString(CultureInfo.InvariantCulture)}' where MeasurementType='FL' and Unit='ppm';";
            var oleDbCommand = new OleDbCommand(sql, ole);
            var i = oleDbCommand.ExecuteNonQuery();
            return i > 0;
        }

        // 改变所有补正值
        public bool AllChange(string dos)
        {
            var i = 0;
            try
            {
                ole.ConnectionString = Mark + Compensate;
                ole.Open();
                var dosList = dos.Split('_');

                var sql =
                    $"update Setup set strPortFrequencyOffsets='{dosList[0]},{dosList[1]},{dosList[2]},{dosList[3]}';";
                var oleDbCommand = new OleDbCommand(sql, ole);
                i = oleDbCommand.ExecuteNonQuery();
            }
            catch (Exception)
            {
                ole.Close();
                return false;
            }

            ole.Close();
            return i > 0;
        }

        // 获取标品值
        public List<StandardData> GetStandard(dynamic data)
        {
            try
            {
                var path = (string)data.path;
                var pn = (string)data.pn;
                var location = (string)data.location;
                var password = (string)data.password;
                ole.ConnectionString = Mark + path + StandardArticle + password;
                ole.Open();
                var inst = new OleDbDataAdapter($"SELECT * FROM GoldenSampleDataList where PN='{pn}' and Location='{location}'", ole);
                var dataTable = new DataTable();
                inst.Fill(dataTable);
                ole.Close();
                if (dataTable.Rows.Count == 0) return null;
                var datas = new List<StandardData>();
                for (var i = 0; i < dataTable.Rows.Count; i++)
                {
                    var num = double.Parse(dataTable.Rows[i][11].ToString()).ToString("F");
                    datas.Add(new StandardData(dataTable.Rows[i][1].ToString(), num));
                }

                dataTable.Reset();
                return datas;
            }
            catch (Exception)
            {
                ole.Close();
                return null;
            }
        }

        // 获取测试值设定
        public List<int> GetTestSet()
        {
            try
            {
                ole.ConnectionString = Mark + Compensate;
                ole.Open();
                var inst = new OleDbDataAdapter("SELECT HighLimit,LowLimit FROM Tests where MeasurementType='FL' and Unit='ppm'", ole);
                var dataSet = new DataTable();
                inst.Fill(dataSet);
                var datas = new List<int>();
                for (var i = 0; i < 2; i++)
                {
                    datas.Add(int.Parse(dataSet.Rows[0][i].ToString()));
                }

                ole.Close();
                dataSet.Reset();
                return datas;
            }
            catch
            {
                ole.Close();
                return null;
            }
        }

        // 初始化
        public void Init()
        {
            ole = new OleDbConnection();
        }
    }
}