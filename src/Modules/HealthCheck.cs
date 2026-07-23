using System;

namespace SarwarSmartPCMS.Modules
{
    internal static class HealthCheck
    {
        public static void Run()
        {
            dynamic ws=ExcelHelpers.Sheet; int id=ExcelHelpers.FindColumn(ws,"task_code","Activity ID","Activity Id"); int name=ExcelHelpers.FindColumn(ws,"task_name","Activity Name"); int dur=ExcelHelpers.FindColumn(ws,"target_drtn_hr_cnt","Original Duration"); int tf=ExcelHelpers.FindColumn(ws,"total_float_hr_cnt","Total Float"); int pred=ExcelHelpers.FindColumn(ws,"Predecessors","Predecessor"); int succ=ExcelHelpers.FindColumn(ws,"Successors","Successor"); if(id==0)throw new Exception("Activity ID column was not found.");
            dynamic outWs=ExcelHelpers.GetOrCreateSheet("Schedule Health Check");outWs.Cells.Clear();outWs.Range("A1:F1").Value2=new object[,]{{"Check","Activity ID","Activity Name","Value","Severity","Recommendation"}};ExcelHelpers.Header(outWs.Range("A1:F1"));int row=2,last=ExcelHelpers.LastRow(ws,id),count=0;
            for(int r=2;r<=last;r++){string aid=Convert.ToString(ws.Cells(r,id).Value2),an=name>0?Convert.ToString(ws.Cells(r,name).Value2):"";double v;
                if(dur>0&&double.TryParse(Convert.ToString(ws.Cells(r,dur).Value2),out v)&&v>480){Add(outWs,ref row,"Long duration",aid,an,v,"Medium","Break into manageable activities.");count++;}
                if(tf>0&&double.TryParse(Convert.ToString(ws.Cells(r,tf).Value2),out v)&&v<0){Add(outWs,ref row,"Negative float",aid,an,v,"High","Review constraints and critical-path logic.");count++;}
                if(pred>0&&string.IsNullOrWhiteSpace(Convert.ToString(ws.Cells(r,pred).Value2))){Add(outWs,ref row,"Open start",aid,an,"","Medium","Add a logical predecessor unless it is a start milestone.");count++;}
                if(succ>0&&string.IsNullOrWhiteSpace(Convert.ToString(ws.Cells(r,succ).Value2))){Add(outWs,ref row,"Open finish",aid,an,"","Medium","Add a logical successor unless it is a finish milestone.");count++;}
            }
            outWs.Cells(1,8).Value2="Issues";outWs.Cells(2,8).Value2=count;outWs.Range("H1:H2").Font.Bold=true;ExcelHelpers.AutoFit(outWs);outWs.Activate();
        }
        private static void Add(dynamic ws,ref int r,string check,string id,string name,object value,string severity,string recommendation){ws.Cells(r,1).Value2=check;ws.Cells(r,2).Value2=id;ws.Cells(r,3).Value2=name;ws.Cells(r,4).Value2=value;ws.Cells(r,5).Value2=severity;ws.Cells(r,6).Value2=recommendation;if(severity=="High")ws.Range(ws.Cells(r,1),ws.Cells(r,6)).Interior.Color=0xC7C7FF;r++;}
    }
}
