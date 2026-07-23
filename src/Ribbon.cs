using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Runtime.InteropServices;

namespace SarwarSmartPCMS
{
    [ComVisible(true)]
    public sealed class Ribbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
 <ribbon>
  <tabs>
   <tab id='tabSarwarSmartPCMS' label='Sarwar Smart PCMS' insertAfterMso='TabHelp'>
    <group id='grpWbs' label='WBS'>
     <button id='btnWbsColor' label='WBS Coloring' imageMso='CellFillColorPicker' size='large' onAction='WbsColoring'/>
     <button id='btnWbsSummary' label='WBS Summary' imageMso='PivotTableInsert' onAction='WbsSummary'/>
     <button id='btnGrouping' label='Grouping' imageMso='OutlineGroup' onAction='Grouping'/>
     <button id='btnSkipRow' label='Skip 1st Row' imageMso='SelectRow' onAction='ToggleSkipFirstRow'/>
    </group>
    <group id='grpGantt' label='Gantt Chart'>
     <button id='btnGantt' label='Draw Gantt Chart' imageMso='ChartGanttChart' size='large' onAction='DrawGantt'/>
     <button id='btnDeleteBars' label='Delete Bars' imageMso='Delete' onAction='DeleteGantt'/>
     <button id='btnScurve' label='S-Curve' imageMso='ChartInsert' onAction='CreateSCurve'/>
     <button id='btnDistribute' label='Distribute Units' imageMso='TableInsertRowsAbove' onAction='DistributeUnits'/>
    </group>
    <group id='grpXer' label='XER'>
     <button id='btnImportXer' label='Import XER' imageMso='ImportTextFile' size='large' onAction='ImportXer'/>
     <button id='btnExportXer' label='Export XER' imageMso='ExportTextFile' onAction='ExportXer'/>
     <button id='btnExplorer' label='XER Explorer' imageMso='FileOpen' onAction='XerExplorer'/>
     <button id='btnCalendar' label='XER Calendar' imageMso='CalendarInsert' onAction='ExtractCalendars'/>
    </group>
    <group id='grpOps' label='XER Operations'>
     <button id='btnHealth' label='Schedule Check' imageMso='ReviewProtectWorkbook' size='large' onAction='ScheduleCheck'/>
     <button id='btnCompare' label='Schedule Comparison' imageMso='CompareAndMergeWorkbooks' onAction='CompareSchedules'/>
     <button id='btnOos' label='OO-Sequence Solver' imageMso='TraceDependents' onAction='OutOfSequence'/>
     <button id='btnFlow' label='Flow Path' imageMso='DiagramInsert' onAction='FlowPath'/>
     <button id='btnCleanup' label='XER Cleanup' imageMso='ClearFormatting' onAction='CleanupXer'/>
     <button id='btnHalf' label='Half-Step XER' imageMso='SplitCells' onAction='HalfStepXer'/>
    </group>
    <group id='grpTools' label='Tools'>
     <button id='btnCellTools' label='Cell Tools' imageMso='TableToolsDesignTab' size='large' onAction='CellTools'/>
     <button id='btnDashboard' label='Dashboard' imageMso='ChartInsert' onAction='CreateDashboard'/>
     <button id='btnAbout' label='About' imageMso='Info' onAction='About'/>
    </group>
   </tab>
  </tabs>
 </ribbon>
</customUI>";

        public void OnLoad(IRibbonUI ribbonUi) { }
        private static void Run(Action action)
        {
            try { action(); }
            catch (Exception ex) { ExcelDna.Integration.XlCall.Excel(XlCall.xlcAlert, "Sarwar Smart PCMS\n\n" + ex.Message); }
        }

        public void WbsColoring(IRibbonControl c) => Run(Modules.WbsTools.ApplyColoring);
        public void WbsSummary(IRibbonControl c) => Run(Modules.WbsTools.CreateSummary);
        public void Grouping(IRibbonControl c) => Run(Modules.WbsTools.ApplyGrouping);
        public void ToggleSkipFirstRow(IRibbonControl c) => Run(Modules.Settings.ToggleSkipFirstRow);
        public void DrawGantt(IRibbonControl c) => Run(Modules.GanttTools.Draw);
        public void DeleteGantt(IRibbonControl c) => Run(Modules.GanttTools.Delete);
        public void CreateSCurve(IRibbonControl c) => Run(Modules.ProgressTools.CreateSCurve);
        public void DistributeUnits(IRibbonControl c) => Run(Modules.ProgressTools.DistributeUnits);
        public void ImportXer(IRibbonControl c) => Run(Modules.XerTools.Import);
        public void ExportXer(IRibbonControl c) => Run(Modules.XerTools.ExportCurrentWorkbook);
        public void XerExplorer(IRibbonControl c) => Run(Modules.XerTools.Explorer);
        public void ExtractCalendars(IRibbonControl c) => Run(Modules.XerTools.ExtractCalendars);
        public void ScheduleCheck(IRibbonControl c) => Run(Modules.HealthCheck.Run);
        public void CompareSchedules(IRibbonControl c) => Run(Modules.CompareTools.Compare);
        public void OutOfSequence(IRibbonControl c) => Run(Modules.AdvancedTools.OutOfSequence);
        public void FlowPath(IRibbonControl c) => Run(Modules.AdvancedTools.FlowPath);
        public void CleanupXer(IRibbonControl c) => Run(Modules.AdvancedTools.Cleanup);
        public void HalfStepXer(IRibbonControl c) => Run(Modules.AdvancedTools.HalfStep);
        public void CellTools(IRibbonControl c) => Run(Modules.CellTools.ShowMenu);
        public void CreateDashboard(IRibbonControl c) => Run(Modules.DashboardTools.Create);
        public void About(IRibbonControl c) => Run(() => System.Windows.Forms.MessageBox.Show(
            "Sarwar Smart PCMS v1.0\nOriginal Excel Project Controls Add-in\nDesigned for 32-bit Excel.",
            "About Sarwar Smart PCMS"));
    }
}
