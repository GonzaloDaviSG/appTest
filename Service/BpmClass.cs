using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebApplication2.Model;

namespace WebApplication2.Service
{
    public class BpmClass
    {
        public List<DataResponse> getJsonResponse(DataRequest param) {
            List<DataResponse> list = new List<DataResponse>();
            string path = AppDomain.CurrentDomain.BaseDirectory + "json/dataNew.json";
            path = path.Replace("\\bin\\Debug\\netcoreapp3.1", "");
            using (StreamReader jsonStream = File.OpenText(path))
            {
                var json = jsonStream.ReadToEnd();
                list = JsonConvert.DeserializeObject<List<DataResponse>>(json);
            }
            return list;
        }
    }
}
