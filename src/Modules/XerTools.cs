using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SarwarSmartPCMS.Modules
{
    internal static class XerTools
    {
        private static string _lastPath;
        private static Dictionary<string, XerTable> _tables;
        public static void Import()
        {
            string path=ExcelHelpers.PickFile("Primavera XER (*.xer)|*.xer|All files (*.*)|*.*","Import Primavera XER"); if(path==null)return;
            _lastPath=path; _tables=XerParser.Parse(path);
            dynamic wb=ExcelHelpers.App.ActiveWorkbook;
            foreach(var kv in _tables){ dynamic ws=ExcelHelpers.GetOrCreateSheet(SafeName(kv.Key)); ws.Cells.Clear(); WriteTable(ws,kv.Value); }
            System.Windows.Forms.MessageBox.Show(_tables.Count+" XER tables imported.","Sarwar Smart PCMS");
        }
        public static void Explorer()
        {
            if(_tables==null){ Import(); if(_tables==null)return; }
            dynamic ws=ExcelHelpers.GetOrCreateSheet("XER Explorer"); ws.Cells.Clear(); ws.Range("A1:C1").Value2=new object[,]{{"Table","Fields","Records"}}; ExcelHelpers.Header(ws.Range("A1:C1")); int r=2;
            foreach(var t in _tables.Values){ws.Cells(r,1).Value2=t.Name;ws.Cells(r,2).Value2=t.Columns.Count;ws.Cells(r,3).Value2=t.Rows.Count;r++;} ExcelHelpers.AutoFit(ws);ws.Activate();
        }
        public static void ExtractCalendars()
        {
            if(_tables==null){Import();if(_tables==null)return;} XerTable t; if(!_tables.TryGetValue("CALENDAR",out t))throw new Exception("CALENDAR table not found."); dynamic ws=ExcelHelpers.GetOrCreateSheet("Calendars"); ws.Cells.Clear();WriteTable(ws,t);ws.Activate();
        }
        public static void ExportCurrentWorkbook()
        {
            string path=ExcelHelpers.SaveFile("Primavera XER (*.xer)|*.xer","Export XER","SarwarSmartPCMS_Export.xer");if(path==null)return;
            dynamic wb=ExcelHelpers.App.ActiveWorkbook; using(var sw=new StreamWriter(path,false,Encoding.Default)){
                sw.WriteLine("ERMHDR\t8.4\t2026-01-01\tProject\tAdmin\tSarwar Smart PCMS\tUSD");
                foreach(dynamic ws in wb.Worksheets){ if(((string)ws.Name).StartsWith("_")||ws.UsedRange.Rows.Count<2)continue; int lastCol=ExcelHelpers.LastCol(ws,1),lastRow=ExcelHelpers.LastRow(ws,1); sw.WriteLine("%T\t"+ws.Name); var headers=new List<string>();for(int c=1;c<=lastCol;c++)headers.Add(Convert.ToString(ws.Cells(1,c).Value2));sw.WriteLine("%F\t"+string.Join("\t",headers));for(int r=2;r<=lastRow;r++){var vals=new List<string>();for(int c=1;c<=lastCol;c++)vals.Add((Convert.ToString(ws.Cells(r,c).Value2)??"").Replace("\t"," "));sw.WriteLine("%R\t"+string.Join("\t",vals));}}
                sw.WriteLine("%E");
            }
        }
        private static void WriteTable(dynamic ws,XerTable t){int cols=t.Columns.Count;object[,] h=new object[1,cols];for(int c=0;c<cols;c++)h[0,c]=t.Columns[c];ws.Range(ws.Cells(1,1),ws.Cells(1,cols)).Value2=h;ExcelHelpers.Header(ws.Range(ws.Cells(1,1),ws.Cells(1,cols)));if(t.Rows.Count>0){object[,] data=new object[t.Rows.Count,cols];for(int r=0;r<t.Rows.Count;r++)for(int c=0;c<cols;c++)data[r,c]=c<t.Rows[r].Length?t.Rows[r][c]:"";ws.Range(ws.Cells(2,1),ws.Cells(t.Rows.Count+1,cols)).Value2=data;}ExcelHelpers.AutoFit(ws);}
        private static string SafeName(string n){foreach(char c in new[]{':','\\','/','?','*','[',']'})n=n.Replace(c,'_');return n.Length>31?n.Substring(0,31):n;}
    }
}
