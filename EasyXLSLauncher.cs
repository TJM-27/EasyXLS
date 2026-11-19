using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class EasyXlsLauncher
{
    [STAThread]
    private static void Main()
    {
        try
        {
            string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string scriptPath = Path.Combine(appDirectory, "EasyXLS-GUI.ps1");

            if (!File.Exists(scriptPath))
            {
                MessageBox.Show(
                    "EasyXLS-GUI.ps1 was not found next to EasyXLS.exe.",
                    "EasyXLS",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -STA -ExecutionPolicy Bypass -File \"" + scriptPath + "\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                WorkingDirectory = appDirectory
            };

            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Unable to start EasyXLS:" + Environment.NewLine + ex.Message,
                "EasyXLS",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
