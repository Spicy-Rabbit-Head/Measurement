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
        public bool AllChange(object dos)
        {
            var i = 0;
            try
            {
                ole.Open();
                var dosList = (List<string>)dos;
                var sql =
                    $"update Setup set strPortFrequencyOffsets='{dosList[0]},{dosList[1]},{dosList[2]},{dosList[3]}';";
                var oleDbCommand = new OleDbCommand(sql, ole);
                i = oleDbCommand.ExecuteNonQuery();
                ole.Close();
            }
            catch (Exception e)
            {
                ole.Close();
                return false;
            }

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

                return datas;
            }
            catch (Exception e)
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

        // 设置为250B连接
        public void Set250B()
        {
            ole.ConnectionString = Mark + Compensate;
        }

        // 设置为标品连接
        public void SetStandard(string s)
        {
            ole.ConnectionString = Mark + s;
        }
    }
}