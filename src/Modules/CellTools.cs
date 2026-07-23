using System;
using System.Windows.Forms;
namespace SarwarSmartPCMS.Modules
{
    internal static class CellTools
    {
        public static void ShowMenu()
        {
            var result=MessageBox.Show("YES: Trim spaces\nNO: Convert selected text to UPPERCASE\nCANCEL: Fill down blanks","Cell Tools",MessageBoxButtons.YesNoCancel);
            dynamic sel=ExcelHelpers.Selection;
            foreach(dynamic cell in sel.Cells){string s=Convert.ToString(cell.Value2);if(result==DialogResult.Yes)cell.Value2=s?.Trim();else if(result==DialogResult.No)cell.Value2=s?.ToUpperInvariant();else if(result==DialogResult.Cancel&&string.IsNullOrWhiteSpace(s)&&cell.Row>sel.Row)cell.Value2=cell.Offset(-1,0).Value2;}
        }
    }
}
