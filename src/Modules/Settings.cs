namespace SarwarSmartPCMS.Modules
{
    internal static class Settings
    {
        public static bool SkipFirstRow { get; private set; } = true;
        public static void ToggleSkipFirstRow()
        {
            SkipFirstRow = !SkipFirstRow;
            System.Windows.Forms.MessageBox.Show("Skip first row: " + (SkipFirstRow ? "ON" : "OFF"), "Sarwar Smart PCMS");
        }
    }
}
