using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace Synergique_Activity_Formatter.Core
{
    public class JsonManager
    {
        public List<Item> ReadData(string fileName)
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appDataFolder = Path.Combine(localAppData, "ActivityFormatter");

            if (!Directory.Exists(appDataFolder))
            {
                Directory.CreateDirectory(appDataFolder);
            }

            return JsonSerializer.Deserialize<List<Item>>(File.ReadAllText(appDataFolder + "\\"+fileName));
        }

        public void SerializeToJson(List<Item> items, string fileName)
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string appDataFolder = Path.Combine(localAppData, "ActivityFormatter");

            if (!Directory.Exists(appDataFolder))
            {
                Directory.CreateDirectory(appDataFolder);
            }

            string jsonString = JsonSerializer.Serialize(items);
            // Console.WriteLine(jsonString);
            File.WriteAllText(appDataFolder + "\\" + fileName, jsonString);
        }
    }
}