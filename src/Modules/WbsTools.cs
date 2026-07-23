using System;
using System.Collections.Generic;
using System.Linq;

namespace SarwarSmartPCMS.Modules
{
    internal static class WbsTools
    {
        private static readonly int[] Palette =
        {
            0x7D4C2E,
            0xA9D18E,
            0xD9EAD3,
            0xDDEBF7,
            0xE2F0D9,
            0xFFF2CC,
            0xFCE4D6,
            0xE4DFEC
        };

        public static void ApplyColoring()
        {
            dynamic ws =
                ExcelHelpers.Sheet;

            dynamic selection =
                ExcelHelpers.Selection;

            int firstRow =
                selection.Row +
                (Settings.SkipFirstRow ? 1 : 0);

            int lastRow =
                selection.Row +
                selection.Rows.Count -
                1;

            int selectedColumn =
                selection.Column;

            for (
                int rowNumber = firstRow;
                rowNumber <= lastRow;
                rowNumber++
            )
            {
                string text =
                    Convert.ToString(
                        ws.Cells(
                            rowNumber,
                            selectedColumn
                        ).Value2
                    ) ?? "";

                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                int level =
                    Math.Max(
                        0,
                        text.Split('.').Length - 1
                    );

                level =
                    Math.Min(
                        level,
                        Palette.Length - 1
                    );

                dynamic rowRange =
                    ws.Range(
                        ws.Cells(
                            rowNumber,
                            selection.Column
                        ),
                        ws.Cells(
                            rowNumber,
                            selection.Column +
                            selection.Columns.Count -
                            1
                        )
                    );

                rowRange.Interior.Color =
                    Palette[level];

                rowRange.Font.Bold =
                    level <= 2;
            }
        }

        public static void ApplyGrouping()
        {
            dynamic ws =
                ExcelHelpers.Sheet;

            dynamic selection =
                ExcelHelpers.Selection;

            int firstRow =
                selection.Row +
                (Settings.SkipFirstRow ? 1 : 0);

            int lastRow =
                selection.Row +
                selection.Rows.Count -
                1;

            int selectedColumn =
                selection.Column;

            ws.Cells.ClearOutline();

            var groupStartRows =
                new Dictionary<int, int>();

            int previousLevel = 0;

            for (
                int rowNumber = firstRow;
                rowNumber <= lastRow + 1;
                rowNumber++
            )
            {
                int currentLevel = 0;

                if (rowNumber <= lastRow)
                {
                    string text =
                        Convert.ToString(
                            ws.Cells(
                                rowNumber,
                                selectedColumn
                            ).Value2
                        ) ?? "";

                    currentLevel =
                        string.IsNullOrWhiteSpace(text)
                            ? previousLevel
                            : Math.Min(
                                8,
                                text.Split('.').Length
                            );
                }

                if (currentLevel > previousLevel)
                {
                    for (
                        int level =
                            previousLevel + 1;
                        level <= currentLevel;
                        level++
                    )
                    {
                        groupStartRows[level] =
                            rowNumber;
                    }
                }

                if (currentLevel < previousLevel)
                {
                    for (
                        int level = previousLevel;
                        level > currentLevel;
                        level--
                    )
                    {
                        if (groupStartRows.ContainsKey(level) &&
                            rowNumber - 1 >=
                            groupStartRows[level])
                        {
                            ws.Rows(
                                groupStartRows[level] +
                                ":" +
                                (rowNumber - 1)
                            ).Rows.Group();
                        }
                    }
                }

                previousLevel =
                    currentLevel;
            }

            ws.Outline.ShowLevels(
                RowLevels: 2
            );
        }

        public static void CreateSummary()
        {
            dynamic ws =
                ExcelHelpers.Sheet;

            dynamic selection =
                ExcelHelpers.Selection;

            dynamic outputSheet =
                ExcelHelpers.GetOrCreateSheet(
                    "WBS Summary"
                );

            outputSheet.Cells.Clear();

            outputSheet.Range(
                "A1:E1"
            ).Value2 = new object[,]
            {
                {
                    "WBS",
                    "Count",
                    "Minimum",
                    "Maximum",
                    "Sum"
                }
            };

            ExcelHelpers.Header(
                outputSheet.Range("A1:E1")
            );

            var summaryData =
                new Dictionary<
                    string,
                    List<double>
                >();

            int startRow =
                selection.Row +
                (Settings.SkipFirstRow ? 1 : 0);

            int endRow =
                selection.Row +
                selection.Rows.Count -
                1;

            for (
                int rowNumber = startRow;
                rowNumber <= endRow;
                rowNumber++
            )
            {
                string wbs =
                    Convert.ToString(
                        ws.Cells(
                            rowNumber,
                            selection.Column
                        ).Value2
                    );

                if (string.IsNullOrWhiteSpace(wbs))
                {
                    continue;
                }

                double value = 0;

                double.TryParse(
                    Convert.ToString(
                        ws.Cells(
                            rowNumber,
                            selection.Column + 1
                        ).Value2
                    ),
                    out value
                );

                if (!summaryData.ContainsKey(wbs))
                {
                    summaryData[wbs] =
                        new List<double>();
                }

                summaryData[wbs].Add(value);
            }

            int outputRow = 2;

            foreach (
                KeyValuePair<
                    string,
                    List<double>
                > item in summaryData
            )
            {
                outputSheet.Cells(
                    outputRow,
                    1
                ).Value2 = item.Key;

                outputSheet.Cells(
                    outputRow,
                    2
                ).Value2 = item.Value.Count;

                outputSheet.Cells(
                    outputRow,
                    3
                ).Value2 =
                    item.Value.Count > 0
                        ? item.Value.Min()
                        : 0;

                outputSheet.Cells(
                    outputRow,
                    4
                ).Value2 =
                    item.Value.Count > 0
                        ? item.Value.Max()
                        : 0;

                outputSheet.Cells(
                    outputRow,
                    5
                ).Value2 =
                    item.Value.Sum();

                outputRow++;
            }

            ExcelHelpers.AutoFit(
                outputSheet
            );

            outputSheet.Activate();
        }
    }
}
