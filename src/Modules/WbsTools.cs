using System;
using System.Collections.Generic;

namespace SarwarSmartPCMS.Modules
{
    internal static class WbsTools
    {
        private static readonly int[] Palette = { 0x7D4C2E, 0xA9D18E, 0xD9EAD3, 0xDDEBF7, 0xE2F0D9, 0xFFF2CC, 0xFCE4D6, 0xE4DFEC };
        public static void ApplyColoring()
        {
            dynamic ws = ExcelHelpers.Sheet;
            dynamic sel = ExcelHelpers.Selection;
            int first = sel.Row + (Settings.SkipFirstRow ? 1 : 0);
            int last = sel.Row + sel.Rows.Count - 1;
            int col = sel.Column;
            for (int r = first; r <= last; r++)
            {
                string text = Convert.ToString(ws.Cells(r, col).Value2) ?? "";
                if (string.IsNullOrWhiteSpace(text)) continue;
                int level = Math.Max(0, text.Split('.').Length - 1);
                level = Math.Min(level, Palette.Length - 1);
                dynamic row = ws.Range(ws.Cells(r, sel.Column), ws.Cells(r, sel.Column + sel.Columns.Count - 1));
                row.Interior.Color = Palette[level];
                row.Font.Bold = level <= 2;
            }
        }
        public static void ApplyGrouping()
        {
            dynamic ws = ExcelHelpers.Sheet;
            dynamic sel = ExcelHelpers.Selection;
            int first = sel.Row + (Settings.SkipFirstRow ? 1 : 0);
            int last = sel.Row + sel.Rows.Count - 1;
            int col = sel.Column;
            ws.Cells.ClearOutline();
            var starts = new Dictionary<int, int>();
            int previousLevel = 0;
            for (int r = first; r <= last + 1; r++)
            {
                int level = 0;
                if (r <= last)
                {
                    string t = Convert.ToString(ws.Cells(r, col).Value2) ?? "";
                    level = string.IsNullOrWhiteSpace(t) ? previousLevel : Math.Min(8, t.Split('.').Length);
                }
                if (level > previousLevel)
                    for (int l = previousLevel + 1; l <= level; l++) starts[l] = r;
                if (level < previousLevel)
                    for (int l = previousLevel; l > level; l--)
                        if (starts.ContainsKey(l) && r - 1 >= starts[l]) ws.Rows(starts[l] + ":" + (r - 1)).Rows.Group();
                previousLevel = level;
            }
            ws.Outline.ShowLevels(RowLevels: 2);
        }
        public static void CreateSummary()
        {
            dynamic ws = ExcelHelpers.Sheet;
            dynamic sel = ExcelHelpers.Selection;
            dynamic outWs = ExcelHelpers.GetOrCreateSheet("WBS Summary");
            outWs.Cells.Clear();
            outWs.Range("A1:E1").Value2 = new object[,] { { "WBS", "Count", "Minimum", "Maximum", "Sum" } };
            ExcelHelpers.Header(outWs.Range("A1:E1"));
            var map = new Dictionary<string, List<double>>();
            int start = sel.Row + (Settings.SkipFirstRow ? 1 : 0);
            int end = sel.Row + sel.Rows.Count - 1;
            for (int r = start; r <= end; r++)
            {
                string key = Convert.ToString(ws.Cells(r, sel.Column).Value2);
                if (string.IsNullOrWhiteSpace(key)) continue;
                double val = 0; double.TryParse(Convert.ToString(ws.Cells(r, sel.Column + 1).Value2), out val);
                if (!map.ContainsKey(key)) map[key] = new List<double>();
                map[key].Add(val);
            }
            int row = 2;
            foreach (var kv in map)
            {
                outWs.Cells(row, 1).Value2 = kv.Key;
                outWs.Cells(row, 2).Value2 = kv.Value.Count;
                outWs.Cells(row, 3).Value2 = kv.Value.Count > 0 ? Math.Min(kv.Value.ToArray()) : 0;
                outWs.Cells(row, 4).Value2 = kv.Value.Count > 0 ? Math.Max(kv.Value.ToArray()) : 0;
                double sum = 0; foreach (double v in kv.Value) sum += v;
                outWs.Cells(row, 5).Value2 = sum;
                row++;
            }
            ExcelHelpers.AutoFit(outWs);
            outWs.Activate();
        }
    }
}
