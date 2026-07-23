using System;

namespace SarwarSmartPCMS.Modules
{
    internal static class ProgressTools
    {
        public static void DistributeUnits()
        {
            dynamic ws = ExcelHelpers.Sheet;

            int startCol = ExcelHelpers.FindColumn(
                ws,
                "Start",
                "Start Date"
            );

            int finishCol = ExcelHelpers.FindColumn(
                ws,
                "Finish",
                "Finish Date"
            );

            int unitsCol = ExcelHelpers.FindColumn(
                ws,
                "Units",
                "Budgeted Units",
                "Quantity"
            );

            if (startCol == 0 ||
                finishCol == 0 ||
                unitsCol == 0)
            {
                throw new Exception(
                    "Start, Finish and Units/Quantity headers are required."
                );
            }

            int lastRow =
                ExcelHelpers.LastRow(ws, 1);

            int outputStartColumn =
                ExcelHelpers.LastCol(ws, 1) + 2;

            DateTime minimumDate =
                DateTime.MaxValue;

            DateTime maximumDate =
                DateTime.MinValue;

            for (int r = 2; r <= lastRow; r++)
            {
                DateTime startDate =
                    DateTime.MinValue;

                DateTime finishDate =
                    DateTime.MinValue;

                if (DateTime.TryParse(
                        Convert.ToString(
                            ws.Cells(r, startCol).Text
                        ),
                        out startDate
                    ) &&
                    startDate < minimumDate)
                {
                    minimumDate = startDate;
                }

                if (DateTime.TryParse(
                        Convert.ToString(
                            ws.Cells(r, finishCol).Text
                        ),
                        out finishDate
                    ) &&
                    finishDate > maximumDate)
                {
                    maximumDate = finishDate;
                }
            }

            if (minimumDate == DateTime.MaxValue)
            {
                throw new Exception(
                    "No valid dates found."
                );
            }

            DateTime currentMonth =
                new DateTime(
                    minimumDate.Year,
                    minimumDate.Month,
                    1
                );

            int currentColumn =
                outputStartColumn;

            while (currentMonth <= maximumDate)
            {
                ws.Cells(
                    1,
                    currentColumn
                ).Value2 = currentMonth;

                ws.Cells(
                    1,
                    currentColumn
                ).NumberFormat = "mmm-yy";

                currentMonth =
                    currentMonth.AddMonths(1);

                currentColumn++;
            }

            for (int r = 2; r <= lastRow; r++)
            {
                DateTime startDate =
                    DateTime.MinValue;

                DateTime finishDate =
                    DateTime.MinValue;

                double units = 0;

                bool validStart =
                    DateTime.TryParse(
                        Convert.ToString(
                            ws.Cells(r, startCol).Text
                        ),
                        out startDate
                    );

                bool validFinish =
                    DateTime.TryParse(
                        Convert.ToString(
                            ws.Cells(r, finishCol).Text
                        ),
                        out finishDate
                    );

                bool validUnits =
                    double.TryParse(
                        Convert.ToString(
                            ws.Cells(r, unitsCol).Value2
                        ),
                        out units
                    );

                if (!validStart ||
                    !validFinish ||
                    !validUnits)
                {
                    continue;
                }

                int numberOfMonths =
                    ((finishDate.Year - startDate.Year) * 12)
                    + finishDate.Month
                    - startDate.Month
                    + 1;

                numberOfMonths =
                    Math.Max(1, numberOfMonths);

                double monthlyUnits =
                    units / numberOfMonths;

                for (
                    int c = outputStartColumn;
                    c < currentColumn;
                    c++
                )
                {
                    object headerValue =
                        ws.Cells(1, c).Value2;

                    if (headerValue == null)
                    {
                        continue;
                    }

                    double oaDate = 0;

                    if (!double.TryParse(
                            Convert.ToString(headerValue),
                            out oaDate
                        ))
                    {
                        continue;
                    }

                    DateTime monthStart =
                        DateTime.FromOADate(oaDate);

                    DateTime monthFinish =
                        monthStart
                            .AddMonths(1)
                            .AddDays(-1);

                    if (monthStart <= finishDate &&
                        monthFinish >= startDate)
                    {
                        ws.Cells(
                            r,
                            c
                        ).Value2 = monthlyUnits;
                    }
                }
            }
        }

        public static void CreateSCurve()
        {
            dynamic ws = ExcelHelpers.Sheet;

            int lastCol =
                ExcelHelpers.LastCol(ws, 1);

            int lastRow =
                ExcelHelpers.LastRow(ws, 1);

            dynamic outWs =
                ExcelHelpers.GetOrCreateSheet(
                    "S-Curve"
                );

            outWs.Cells.Clear();

            outWs.Range(
                "A1:C1"
            ).Value2 = new object[,]
            {
                {
                    "Period",
                    "Monthly",
                    "Cumulative"
                }
            };

            ExcelHelpers.Header(
                outWs.Range("A1:C1")
            );

            int outputRow = 2;
            double cumulative = 0;

            for (int c = 1; c <= lastCol; c++)
            {
                DateTime periodDate =
                    DateTime.MinValue;

                if (!DateTime.TryParse(
                        Convert.ToString(
                            ws.Cells(1, c).Text
                        ),
                        out periodDate
                    ))
                {
                    continue;
                }

                double monthlyTotal = 0;

                for (int r = 2; r <= lastRow; r++)
                {
                    double value = 0;

                    if (double.TryParse(
                            Convert.ToString(
                                ws.Cells(r, c).Value2
                            ),
                            out value
                        ))
                    {
                        monthlyTotal += value;
                    }
                }

                cumulative += monthlyTotal;

                outWs.Cells(
                    outputRow,
                    1
                ).Value2 = periodDate;

                outWs.Cells(
                    outputRow,
                    1
                ).NumberFormat = "mmm-yy";

                outWs.Cells(
                    outputRow,
                    2
                ).Value2 = monthlyTotal;

                outWs.Cells(
                    outputRow,
                    3
                ).Value2 = cumulative;

                outputRow++;
            }

            if (outputRow <= 2)
            {
                throw new Exception(
                    "No distributed period data found."
                );
            }

            dynamic chart =
                outWs.Shapes
                    .AddChart2(
                        227,
                        65,
                        300,
                        20,
                        700,
                        360
                    )
                    .Chart;

            chart.SetSourceData(
                outWs.Range(
                    "A1:C" + (outputRow - 1)
                )
            );

            chart.ChartType = 4;
            chart.HasTitle = true;
            chart.ChartTitle.Text =
                "Sarwar Smart PCMS S-Curve";

            ExcelHelpers.AutoFit(outWs);
            outWs.Activate();
        }
    }
}
