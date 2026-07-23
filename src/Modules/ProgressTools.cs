using System;

namespace SarwarSmartPCMS.Modules
{
    internal static class ProgressTools
    {
        public static void DistributeUnits()
        {
            dynamic ws = ExcelHelpers.Sheet;
            int startCol = ExcelHelpers.FindColumn(ws, "Start", "Start Date");
            int finishCol = ExcelHelpers.FindColumn(ws, "Finish", "Finish Date");
            int unitsCol = ExcelHelpers.FindColumn(ws, "Units", "Budgeted Units", "Quantity");
            if (startCol == 0 || finishCol == 0 || unitsCol == 0) throw new Exception("Start, Finish and Units/Quantity headers are required.");
            int last = ExcelHelpers.LastRow(ws, 1);
            int output = ExcelHelpers.LastCol(ws, 1) + 2;
            DateTime min = DateTime.MaxValue, max = DateTime.MinValue;
            for (int r = 2; r <= last; r++) { DateTime s,f; if(DateTime.TryParse(Convert.ToString(ws.Cells(r,startCol).Text),out s)&&s<min)min=s; if(DateTime.TryParse(Convert.ToString(ws.Cells(r,finishCol).Text),out f)&&f>max)max=f; }
            if(min==DateTime.MaxValue) throw new Exception("No valid dates found.");
            DateTime cur = new DateTime(min.Year,min.Month,1); int c=output;
            while(cur<=max){ ws.Cells(1,c).Value2=cur; ws.Cells(1,c).NumberFormat="mmm-yy"; cur=cur.AddMonths(1); c++; }
            for(int r=2;r<=last;r++){
                DateTime s,f; double units;
                if(!DateTime.TryParse(Convert.ToString(ws.Cells(r,startCol).Text),out s)||!DateTime.TryParse(Convert.ToString(ws.Cells(r,finishCol).Text),out f)||!double.TryParse(Convert.ToString(ws.Cells(r,unitsCol).Value2),out units))continue;
                int months=Math.Max(1,((f.Year-s.Year)*12+f.Month-s.Month)+1); double each=units/months;
                for(int cc=output;cc<c;cc++){ DateTime m=DateTime.FromOADate(Convert.ToDouble(ws.Cells(1,cc).Value2)); if(m<=f&&m.AddMonths(1).AddDays(-1)>=s) ws.Cells(r,cc).Value2=each; }
            }
        }
        public static void CreateSCurve()
        {
            dynamic ws = ExcelHelpers.Sheet;
            int lastCol = ExcelHelpers.LastCol(ws, 1);
            int lastRow = ExcelHelpers.LastRow(ws, 1);
            dynamic outWs = ExcelHelpers.GetOrCreateSheet("S-Curve"); outWs.Cells.Clear();
            outWs.Range("A1:C1").Value2 = new object[,]{{"Period","Monthly","Cumulative"}}; ExcelHelpers.Header(outWs.Range("A1:C1"));
            int row=2; double cumulative=0;
            for(int c=1;c<=lastCol;c++){
                DateTime d; if(!DateTime.TryParse(Convert.ToString(ws.Cells(1,c).Text),out d)) continue;
                double sum=0; for(int r=2;r<=lastRow;r++){ double v; if(double.TryParse(Convert.ToString(ws.Cells(r,c).Value2),out v)) sum+=v; }
                cumulative+=sum; outWs.Cells(row,1).Value2=d; outWs.Cells(row,1).NumberFormat="mmm-yy"; outWs.Cells(row,2).Value2=sum; outWs.Cells(row,3).Value2=cumulative; row++;
            }
            if(row<=2) throw new Exception("No distributed period data found.");
            dynamic chart=outWs.Shapes.AddChart2(227,65,300,20,700,360).Chart;
            chart.SetSourceData(outWs.Range("A1:C"+(row-1))); chart.ChartType=4; chart.HasTitle=true; chart.ChartTitle.Text="Sarwar Smart PCMS S-Curve";
            ExcelHelpers.AutoFit(outWs); outWs.Activate();
        }
    }
}
