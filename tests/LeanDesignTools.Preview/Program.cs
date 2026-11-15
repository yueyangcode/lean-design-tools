using System;
using System.Drawing;
using System.Windows.Forms;
using LeanDesignTools;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new Form
        {
            Text = "设计工具预览",
            ClientSize = new Size(220, 790),
            StartPosition = FormStartPosition.CenterScreen,
            Controls = { new DesignPane { Dock = DockStyle.Fill } }
        });
    }
}
