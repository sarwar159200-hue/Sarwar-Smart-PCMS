using System;
namespace SarwarSmartPCMS.Modules
{
    internal static class AdvancedTools
    {
        private static void Notice(string feature)=>System.Windows.Forms.MessageBox.Show(feature+" foundation is included. Full production logic is scheduled for the next release after testing with real project files.","Sarwar Smart PCMS");
        public static void OutOfSequence()=>Notice("Out-of-sequence solver");
        public static void FlowPath()=>Notice("Flow-path analysis");
        public static void Cleanup()=>Notice("XER cleanup");
        public static void HalfStep()=>Notice("Half-step XER");
    }
}
