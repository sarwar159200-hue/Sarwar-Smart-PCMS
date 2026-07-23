using System;

namespace SarwarSmartPCMS.Modules
{
    internal static class HealthCheck
    {
        public static void Run()
        {
            dynamic ws =
                ExcelHelpers.Sheet;

            int idColumn =
                ExcelHelpers.FindColumn(
                    ws,
                    "task_code",
                    "Activity ID",
                    "Activity Id"
                );

            int nameColumn =
                ExcelHelpers.FindColumn(
                    ws,
                    "task_name",
                    "Activity Name"
                );

            int durationColumn =
                ExcelHelpers.FindColumn(
                    ws,
                    "target_drtn_hr_cnt",
                    "Original Duration"
                );

            int floatColumn =
                ExcelHelpers.FindColumn(
                    ws,
                    "total_float_hr_cnt",
                    "Total Float"
                );

            int predecessorColumn =
                ExcelHelpers.FindColumn(
                    ws,
                    "Predecessors",
                    "Predecessor"
                );

            int successorColumn =
                ExcelHelpers.FindColumn(
                    ws,
                    "Successors",
                    "Successor"
                );

            if (idColumn == 0)
            {
                throw new Exception(
                    "Activity ID column was not found."
                );
            }

            dynamic outputSheet =
                ExcelHelpers.GetOrCreateSheet(
                    "Schedule Health Check"
                );

            outputSheet.Cells.Clear();

            outputSheet.Range(
                "A1:F1"
            ).Value2 = new object[,]
            {
                {
                    "Check",
                    "Activity ID",
                    "Activity Name",
                    "Value",
                    "Severity",
                    "Recommendation"
                }
            };

            ExcelHelpers.Header(
                outputSheet.Range("A1:F1")
            );

            int outputRow = 2;

            int lastRow =
                ExcelHelpers.LastRow(
                    ws,
                    idColumn
                );

            int issueCount = 0;

            for (int r = 2; r <= lastRow; r++)
            {
                string activityId =
                    Convert.ToString(
                        ws.Cells(
                            r,
                            idColumn
                        ).Value2
                    );

                string activityName =
                    nameColumn > 0
                        ? Convert.ToString(
                            ws.Cells(
                                r,
                                nameColumn
                            ).Value2
                        )
                        : "";

                double value = 0;

                if (durationColumn > 0 &&
                    double.TryParse(
                        Convert.ToString(
                            ws.Cells(
                                r,
                                durationColumn
                            ).Value2
                        ),
                        out value
                    ) &&
                    value > 480)
                {
                    AddIssue(
                        outputSheet,
                        ref outputRow,
                        "Long duration",
                        activityId,
                        activityName,
                        value,
                        "Medium",
                        "Break into manageable activities."
                    );

                    issueCount++;
                }

                value = 0;

                if (floatColumn > 0 &&
                    double.TryParse(
                        Convert.ToString(
                            ws.Cells(
                                r,
                                floatColumn
                            ).Value2
                        ),
                        out value
                    ) &&
                    value < 0)
                {
                    AddIssue(
                        outputSheet,
                        ref outputRow,
                        "Negative float",
                        activityId,
                        activityName,
                        value,
                        "High",
                        "Review constraints and critical-path logic."
                    );

                    issueCount++;
                }

                if (predecessorColumn > 0 &&
                    string.IsNullOrWhiteSpace(
                        Convert.ToString(
                            ws.Cells(
                                r,
                                predecessorColumn
                            ).Value2
                        )
                    ))
                {
                    AddIssue(
                        outputSheet,
                        ref outputRow,
                        "Open start",
                        activityId,
                        activityName,
                        "",
                        "Medium",
                        "Add a logical predecessor unless it is a start milestone."
                    );

                    issueCount++;
                }

                if (successorColumn > 0 &&
                    string.IsNullOrWhiteSpace(
                        Convert.ToString(
                            ws.Cells(
                                r,
                                successorColumn
                            ).Value2
                        )
                    ))
                {
                    AddIssue(
                        outputSheet,
                        ref outputRow,
                        "Open finish",
                        activityId,
                        activityName,
                        "",
                        "Medium",
                        "Add a logical successor unless it is a finish milestone."
                    );

                    issueCount++;
                }
            }

            outputSheet.Cells(
                1,
                8
            ).Value2 = "Issues";

            outputSheet.Cells(
                2,
                8
            ).Value2 = issueCount;

            outputSheet.Range(
                "H1:H2"
            ).Font.Bold = true;

            ExcelHelpers.AutoFit(
                outputSheet
            );

            outputSheet.Activate();
        }

        private static void AddIssue(
            dynamic ws,
            ref int row,
            string check,
            string activityId,
            string activityName,
            object value,
            string severity,
            string recommendation
        )
        {
            ws.Cells(row, 1).Value2 =
                check;

            ws.Cells(row, 2).Value2 =
                activityId;

            ws.Cells(row, 3).Value2 =
                activityName;

            ws.Cells(row, 4).Value2 =
                value;

            ws.Cells(row, 5).Value2 =
                severity;

            ws.Cells(row, 6).Value2 =
                recommendation;

            if (severity == "High")
            {
                ws.Range(
                    ws.Cells(row, 1),
                    ws.Cells(row, 6)
                ).Interior.Color = 0xC7C7FF;
            }

            row++;
        }
    }
}
