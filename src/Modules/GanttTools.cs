using System;

namespace SarwarSmartPCMS.Modules
{
    internal static class GanttTools
    {
        public static void Draw()
        {
            dynamic ws = ExcelHelpers.Sheet;
            int startCol = ExcelHelpers.FindColumn(ws, "Start", "Start Date", "Planned Start");
            int finishCol = ExcelHelpers.FindColumn(ws, "Finish", "Finish Date", "Planned Finish");
            if (startCol == 0 || finishCol == 0) throw new Exception("Headers 'Start' and 'Finish' were not found in row 1.");
            int lastRow = Math.Max(2, ExcelHelpers.LastRow(ws, 1));
            DateTime min = DateTime.MaxValue, max = DateTime.MinValue;
            for (int r = 2; r <= lastRow; r++)
            {
                DateTime s, f;
                if (DateTime.TryParse(Convert.ToString(ws.Cells(r, startCol).Text), out s) && s < min) min = s;
                if (DateTime.TryParse(Convert.ToString(ws.Cells(r, finishCol).Text), out f) && f > max) max = f;
            }
            if (min == DateTime.MaxValue) throw new Exception("No valid schedule dates were found.");
            int ganttStart = ExcelHelpers.LastCol(ws, 1) + 2;
            DateTime cursor = new DateTime(min.Year, min.Month, 1);
            int c = ganttStart;
            while (cursor <= max)
            {
                ws.Cells(1, c).Value2 = cursor;
                ws.Cells(1, c).NumberFormat = "mmm-yy";
                ws.Cells(1, c).Font.Bold = true;
                ws.Columns(c).ColumnWidth = 4.2;
                for (int r = 2; r <= lastRow; r++)
                {
                    DateTime s, f;
                    if (!DateTime.TryParse(Convert.ToString(ws.Cells(r, startCol).Text), out s) || !DateTime.TryParse(Convert.ToString(ws.Cells(r, finishCol).Text), out f)) continue;
                    DateTime endMonth = cursor.AddMonths(1).AddDays(-1);
                    if (s <= endMonth && f >= cursor) ws.Cells(r, c).Interior.Color = 0x50B000;
                }
                cursor = cursor.AddMonths(1); c++;
            }
            ws.Range(ws.Cells(1, ganttStart), ws.Cells(lastRow, c - 1)).Borders.LineStyle = 1;
        }
        public static void Delete()
        {
            dynamic ws = ExcelHelpers.Sheet;
            int lastCol = ExcelHelpers.LastCol(ws, 1);
            int start = ExcelHelpers.FindColumn(ws, "Gantt Start");
            if (start == 0) start = Math.Max(1, lastCol - 60);
            ws.Range(ws.Cells(1, start), ws.Cells(ExcelHelpers.LastRow(ws, 1), lastCol)).Interior.Pattern = -4142;
        }
    }
}
