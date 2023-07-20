namespace Measurement.Entity
{
    public class StandardData
    {
        public StandardData(string label, string value)
        {
            this.label = label;
            this.value = value;
        }

        public string label { get; set; }
        public string value { get; set; }
    }
}