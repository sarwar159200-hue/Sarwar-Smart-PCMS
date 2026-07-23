using System;

namespace SarwarSmartPCMS.Modules
{
    internal static class GanttTools
    {
        public static void Draw()
        {
            dynamic ws = ExcelHelpers.Sheet;

            int startCol = ExcelHelpers.FindColumn(
                ws,
                "Start",
                "Start Date",
                "Planned Start"
            );

            int finishCol = ExcelHelpers.FindColumn(
                ws,
                "Finish",
                "Finish Date",
                "Planned Finish"
            );

            if (startCol == 0 || finishCol == 0)
            {
                throw new Exception(
                    "Headers 'Start' and 'Finish' were not found in row 1."
                );
            }

            int lastRow = Math.Max(
                2,
                ExcelHelpers.LastRow(ws, 1)
            );

            DateTime min = DateTime.MaxValue;
            DateTime max = DateTime.MinValue;

            for (int r = 2; r <= lastRow; r++)
            {
                DateTime startDate = DateTime.MinValue;
                DateTime finishDate = DateTime.MinValue;

                if (DateTime.TryParse(
                        Convert.ToString(
                            ws.Cells(r, startCol).Text
                        ),
                        out startDate
                    ) &&
                    startDate < min)
                {
                    min = startDate;
                }

                if (DateTime.TryParse(
                        Convert.ToString(
                            ws.Cells(r, finishCol).Text
                        ),
                        out finishDate
                    ) &&
                    finishDate > max)
                {
                    max = finishDate;
                }
            }

            if (min == DateTime.MaxValue)
            {
                throw new Exception(
                    "No valid schedule dates were found."
                );
            }

            int ganttStart =
                ExcelHelpers.LastCol(ws, 1) + 2;

            DateTime cursor =
                new DateTime(min.Year, min.Month, 1);

            int ganttColumn = ganttStart;

            while (cursor <= max)
            {
                ws.Cells(1, ganttColumn).Value2 = cursor;
                ws.Cells(1, ganttColumn).NumberFormat = "mmm-yy";
                ws.Cells(1, ganttColumn).Font.Bold = true;
                ws.Columns(ganttColumn).ColumnWidth = 4.2;

                DateTime endOfMonth =
                    cursor.AddMonths(1).AddDays(-1);

                for (int r = 2; r <= lastRow; r++)
                {
                    DateTime startDate = DateTime.MinValue;
                    DateTime finishDate = DateTime.MinValue;

                    bool hasStart = DateTime.TryParse(
                        Convert.ToString(
                            ws.Cells(r, startCol).Text
                        ),
                        out startDate
                    );

                    bool hasFinish = DateTime.TryParse(
                        Convert.ToString(
                            ws.Cells(r, finishCol).Text
                        ),
                        out finishDate
                    );

                    if (!hasStart || !hasFinish)
                    {
                        continue;
                    }

                    if (startDate <= endOfMonth &&
                        finishDate >= cursor)
                    {
                        ws.Cells(
                            r,
                            ganttColumn
                        ).Interior.Color = 0x50B000;
                    }
                }

                cursor = cursor.AddMonths(1);
                ganttColumn++;
            }

            ws.Range(
                ws.Cells(1, ganttStart),
                ws.Cells(lastRow, ganttColumn - 1)
            ).Borders.LineStyle = 1;
        }

        public static void Delete()
        {
            dynamic ws = ExcelHelpers.Sheet;

            int lastCol =
                ExcelHelpers.LastCol(ws, 1);

            int start =
                ExcelHelpers.FindColumn(
                    ws,
                    "Gantt Start"
                );

            if (start == 0)
            {
                start = Math.Max(
                    1,
                    lastCol - 60
                );
            }

            int lastRow =
                ExcelHelpers.LastRow(ws, 1);

            ws.Range(
                ws.Cells(1, start),
                ws.Cells(lastRow, lastCol)
            ).Interior.Pattern = -4142;
        }
    }
}
