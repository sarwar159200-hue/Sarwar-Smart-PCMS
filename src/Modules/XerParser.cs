using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SarwarSmartPCMS.Modules
{
    internal sealed class XerTable { public string Name; public List<string> Columns = new List<string>(); public List<string[]> Rows = new List<string[]>(); }
    internal static class XerParser
    {
        public static Dictionary<string, XerTable> Parse(string path)
        {
            var result = new Dictionary<string, XerTable>(StringComparer.OrdinalIgnoreCase);
            XerTable current = null;
            foreach (var raw in File.ReadLines(path, Encoding.Default))
            {
                var p = raw.Split('\t'); if (p.Length == 0) continue;
                if (p[0] == "%T" && p.Length > 1) { current = new XerTable { Name = p[1] }; result[current.Name] = current; }
                else if (p[0] == "%F" && current != null) { current.Columns.Clear(); for (int i=1;i<p.Length;i++) current.Columns.Add(p[i]); }
                else if (p[0] == "%R" && current != null) { var row=new string[Math.Max(0,p.Length-1)]; Array.Copy(p,1,row,0,row.Length); current.Rows.Add(row); }
            }
            return result;
        }
    }
}
