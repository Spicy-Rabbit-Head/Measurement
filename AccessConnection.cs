using System.Data.OleDb;
using System.Globalization;

namespace Measurement
{
    internal class AccessConnection
    {
        private OleDbConnection ole;

        public bool Change(double num)
        {
            var sql = $"update Tests set ReferenceValue='{num.ToString(CultureInfo.InvariantCulture)}' where MeasurementType='FL' and Unit='ppm';";
            var oleDbCommand = new OleDbCommand(sql, ole);
            var i = oleDbCommand.ExecuteNonQuery();
            return i > 0;
        }

        public bool AllChange(double[] dos)
        {
            var sql =
                $"update Setup set strPortFrequencyOffsets='{dos[0].ToString(CultureInfo.InvariantCulture)},{dos[1].ToString(CultureInfo.InvariantCulture)},{dos[2].ToString(CultureInfo.InvariantCulture)},{dos[3].ToString(CultureInfo.InvariantCulture)}';";
            var oleDbCommand = new OleDbCommand(sql, ole);
            var i = oleDbCommand.ExecuteNonQuery();
            return i > 0;
        }

        public void Init(double num)
        {
            ole = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\ProgramData\\Saunders & Associates\\250B\\Data\\Server-1.mdb");
        }

        public void Open()
        {
            ole.Open();
        }

        public void Close()
        {
            ole.Close();
        }
    }
}