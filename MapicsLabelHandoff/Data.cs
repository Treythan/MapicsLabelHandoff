using Newtonsoft.Json;

namespace MapicsLabelHandoff
{
    internal class Data
    {
        public string lastReceiver = "Err";
        public string lastPurchaseOrder = "Err";

        public Data LoadData()
        {
            string path = AppContext.BaseDirectory + "data.json";
            if (!File.Exists(path))
                return new Data();
            else
                return JsonConvert.DeserializeObject<Data>(File.ReadAllText(path));
        }
        public void SaveData()
        {
            string newJson = JsonConvert.SerializeObject(this, Formatting.Indented);
            File.WriteAllText(AppContext.BaseDirectory + "data.json", newJson);
        }
    }
}
