using ExcelDna.Integration;
using System;
using System.Collections.Generic;

namespace SarwarSmartPCMS.Modules
{
    internal static class ExcelHelpers
    {
        public static dynamic App => ExcelDnaUtil.Application;
        public static dynamic Sheet => App.ActiveSheet;
        public static dynamic Selection => App.Selection;
        public static int LastRow(dynamic ws, int col = 1) => ws.Cells(ws.Rows.Count, col).End(-4162).Row; // xlUp
        public static int LastCol(dynamic ws, int row = 1) => ws.Cells(row, ws.Columns.Count).End(-4159).Column; // xlToLeft
        public static int FindColumn(dynamic ws, params string[] names)
        {
            int last = LastCol(ws, 1);
            for (int c = 1; c <= last; c++)
            {
                string value = Convert.ToString(ws.Cells(1, c).Value2)?.Trim();
                foreach (var n in names)
                    if (string.Equals(value, n, StringComparison.OrdinalIgnoreCase)) return c;
            }
            return 0;
        }
        public static dynamic GetOrCreateSheet(string name)
        {
            dynamic wb = App.ActiveWorkbook;
            foreach (dynamic ws in wb.Worksheets)
                if (string.Equals((string)ws.Name, name, StringComparison.OrdinalIgnoreCase)) return ws;
            dynamic created = wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
            created.Name = name;
            return created;
        }
        public static void Header(dynamic range)
        {
            range.Font.Bold = true;
            range.Font.Color = 0xFFFFFF;
            range.Interior.Color = 0x2E4C7D;
            range.HorizontalAlignment = -4108;
        }
        public static void AutoFit(dynamic ws) => ws.Cells.EntireColumn.AutoFit();
        public static string PickFile(string filter, string title)
        {
            using (var d = new System.Windows.Forms.OpenFileDialog { Filter = filter, Title = title, Multiselect = false })
                return d.ShowDialog() == System.Windows.Forms.DialogResult.OK ? d.FileName : null;
        }
        public static string SaveFile(string filter, string title, string fileName)
        {
            using (var d = new System.Windows.Forms.SaveFileDialog { Filter = filter, Title = title, FileName = fileName })
                return d.ShowDialog() == System.Windows.Forms.DialogResult.OK ? d.FileName : null;
        }
    }
}
