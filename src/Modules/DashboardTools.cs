using System;
namespace SarwarSmartPCMS.Modules
{
    internal static class DashboardTools
    {
        public static void Create()
        {
            dynamic src=ExcelHelpers.Sheet;dynamic ws=ExcelHelpers.GetOrCreateSheet("PCMS Dashboard");ws.Cells.Clear();ws.Range("A1:F2").Merge();ws.Range("A1").Value2="SARWAR SMART PCMS DASHBOARD";ws.Range("A1").Font.Bold=true;ws.Range("A1").Font.Size=20;ws.Range("A1").Interior.Color=0x2E4C7D;ws.Range("A1").Font.Color=0xFFFFFF;ws.Range("A4:B9").Value2=new object[,]{{"KPI","Value"},{"Activities",Math.Max(0,ExcelHelpers.LastRow(src,1)-1)},{"Completed",CountText(src,"Complete")},{"In Progress",CountText(src,"In Progress")},{"Not Started",CountText(src,"Not Started")},{"Data Date",DateTime.Today}};ExcelHelpers.Header(ws.Range("A4:B4"));ws.Range("B9").NumberFormat="dd-mmm-yyyy";ws.Columns("A:F").ColumnWidth=22;ws.Activate();
        }
        private static int CountText(dynamic ws,string text){int last=ExcelHelpers.LastRow(ws,1),cols=ExcelHelpers.LastCol(ws,1),count=0;for(int r=2;r<=last;r++)for(int c=1;c<=cols;c++)if(string.Equals(Convert.ToString(ws.Cells(r,c).Value2),text,StringComparison.OrdinalIgnoreCase))count++;return count;}
    }
}
