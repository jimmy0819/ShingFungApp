using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShinFung
{
    internal class WordStruct
    {
        public class ShinWord
        {
            public string nameApply { get; set; }
            public string phoneNum { get; set; }
            public string address { get; set; }

            // Names and attributes for 0 to 9
            public string Name0 { get; set; }
            public string zod0 { get; set; }
            public string age0 { get; set; }
            public string v10 { get; set; }
            public string v20 { get; set; }
            public string v30 { get; set; }
            public string sin0 { get; set; }

            public string Name1 { get; set; }
            public string zod1 { get; set; }
            public string age1 { get; set; }
            public string v11 { get; set; }
            public string v21 { get; set; }
            public string v31 { get; set; }
            public string sin1 { get; set; }

            public string Name2 { get; set; }
            public string zod2 { get; set; }
            public string age2 { get; set; }
            public string v12 { get; set; }
            public string v22 { get; set; }
            public string v32 { get; set; }
            public string sin2 { get; set; }

            public string Name3 { get; set; }
            public string zod3 { get; set; }
            public string age3 { get; set; }
            public string v13 { get; set; }
            public string v23 { get; set; }
            public string v33 { get; set; }
            public string sin3 { get; set; }

            public string Name4 { get; set; }
            public string zod4 { get; set; }
            public string age4 { get; set; }
            public string v14 { get; set; }
            public string v24 { get; set; }
            public string v34 { get; set; }
            public string sin4 { get; set; }

            public string Name5 { get; set; }
            public string zod5 { get; set; }
            public string age5 { get; set; }
            public string v15 { get; set; }
            public string v25 { get; set; }
            public string v35 { get; set; }
            public string sin5 { get; set; }

            public string Name6 { get; set; }
            public string zod6 { get; set; }
            public string age6 { get; set; }
            public string v16 { get; set; }
            public string v26 { get; set; }
            public string v36 { get; set; }
            public string sin6 { get; set; }

            public string Name7 { get; set; }
            public string zod7 { get; set; }
            public string age7 { get; set; }
            public string v17 { get; set; }
            public string v27 { get; set; }
            public string v37 { get; set; }
            public string sin7 { get; set; }

            public string Name8 { get; set; }
            public string zod8 { get; set; }
            public string age8 { get; set; }
            public string v18 { get; set; }
            public string v28 { get; set; }
            public string v38 { get; set; }
            public string sin8 { get; set; }

            public string Name9 { get; set; }
            public string zod9 { get; set; }
            public string age9 { get; set; }
            public string v19 { get; set; }
            public string v29 { get; set; }
            public string v39 { get; set; }
            public string sin9 { get; set; }

            // Summary variables
            public string suma { get; set; }
            public string sumb { get; set; }
            public string sumc { get; set; }
            public string sum { get; set; }
            public string oil { get; set; }
            public string bainame { get; set; }
            public string bsum { get; set; }
            public string naname { get; set; }
            public string nsum { get; set; }
            public string employname { get; set; }
            public string year { get; set; }
            public string mon { get; set; }
            public string day { get; set; }

            public ShinWord(Dictionary<string, string> data)
            {
                if (data == null) throw new ArgumentNullException(nameof(data));

                nameApply = data.ContainsKey("nameApply") ? data["nameApply"] : string.Empty;
                phoneNum = data.ContainsKey("phoneNum") ? data["phoneNum"] : string.Empty;
                address = data.ContainsKey("address") ? data["address"] : string.Empty;

                // Person fields for 0 to 9
                for (int i = 0; i < 10; i++)
                {
                    if (data.ContainsKey($"Name{i}")) GetType().GetProperty($"Name{i}")?.SetValue(this, data[$"Name{i}"]);
                    if (data.ContainsKey($"zod{i}")) GetType().GetProperty($"zod{i}")?.SetValue(this, data[$"zod{i}"]);
                    if (data.ContainsKey($"age{i}")) GetType().GetProperty($"age{i}")?.SetValue(this, data[$"age{i}"]);
                    if (data.ContainsKey($"1v{i}")) GetType().GetProperty($"v1{i}")?.SetValue(this, data[$"1v{i}"]);
                    if (data.ContainsKey($"2v{i}")) GetType().GetProperty($"v2{i}")?.SetValue(this, data[$"2v{i}"]);
                    if (data.ContainsKey($"3v{i}")) GetType().GetProperty($"v3{i}")?.SetValue(this, data[$"3v{i}"]);
                    if (data.ContainsKey($"sin{i}")) GetType().GetProperty($"sin{i}")?.SetValue(this, data[$"sin{i}"]);
                }

                // Summary fields
                suma = data.ContainsKey("suma") ? data["suma"] : string.Empty;
                sumb = data.ContainsKey("sumb") ? data["sumb"] : string.Empty;
                sumc = data.ContainsKey("sumc") ? data["sumc"] : string.Empty;
                sum = data.ContainsKey("sum") ? data["sum"] : string.Empty;
                oil = data.ContainsKey("oil") ? data["oil"] : string.Empty;
                bainame = data.ContainsKey("bainame") ? data["bainame"] : string.Empty;
                bsum = data.ContainsKey("bsum") ? data["bsum"] : string.Empty;
                naname = data.ContainsKey("naname") ? data["naname"] : string.Empty;
                nsum = data.ContainsKey("nsum") ? data["nsum"] : string.Empty;
                employname = data.ContainsKey("employname") ? data["employname"] : string.Empty;
                year = data.ContainsKey("year") ? data["year"] : string.Empty;
                mon = data.ContainsKey("mon") ? data["mon"] : string.Empty;
                day = data.ContainsKey("day") ? data["day"] : string.Empty;
            }
        }

    }
}
