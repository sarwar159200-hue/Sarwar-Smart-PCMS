using System;
using System.Collections.Generic;

namespace SarwarSmartPCMS.Modules
{
    internal static class CompareTools
    {
        public static void Compare()
        {
            string a=ExcelHelpers.PickFile("Primavera XER (*.xer)|*.xer","Select previous XER");if(a==null)return;string b=ExcelHelpers.PickFile("Primavera XER (*.xer)|*.xer","Select current XER");if(b==null)return;
            var ta=XerParser.Parse(a);var tb=XerParser.Parse(b);XerTable aa,bb;if(!ta.TryGetValue("TASK",out aa)||!tb.TryGetValue("TASK",out bb))throw new Exception("TASK table missing in one file.");
            int idA=aa.Columns.FindIndex(x=>x.Equals("task_code",StringComparison.OrdinalIgnoreCase));int idB=bb.Columns.FindIndex(x=>x.Equals("task_code",StringComparison.OrdinalIgnoreCase));if(idA<0||idB<0)throw new Exception("task_code field missing.");
            var oldMap=new Dictionary<string,string[]>(StringComparer.OrdinalIgnoreCase);foreach(var r in aa.Rows)if(idA<r.Length)oldMap[r[idA]]=r;
            dynamic ws=ExcelHelpers.GetOrCreateSheet("Schedule Comparison");ws.Cells.Clear();ws.Range("A1:D1").Value2=new object[,]{{"Activity ID","Change","Previous Record","Current Record"}};ExcelHelpers.Header(ws.Range("A1:D1"));int row=2;
            foreach(var r in bb.Rows){if(idB>=r.Length)continue;string id=r[idB];string[] old;if(!oldMap.TryGetValue(id,out old)){Add(ws,ref row,id,"Added","",string.Join(" | ",r));}else if(string.Join("\t",old)!=string.Join("\t",r)){Add(ws,ref row,id,"Modified",string.Join(" | ",old),string.Join(" | ",r));oldMap.Remove(id);}else oldMap.Remove(id);}
            foreach(var kv in oldMap)Add(ws,ref row,kv.Key,"Deleted",string.Join(" | ",kv.Value),"");ExcelHelpers.AutoFit(ws);ws.Activate();
        }
        private static void Add(dynamic ws,ref int row,string id,string change,string old,string cur){ws.Cells(row,1).Value2=id;ws.Cells(row,2).Value2=change;ws.Cells(row,3).Value2=old;ws.Cells(row,4).Value2=cur;row++;}
    }
}
