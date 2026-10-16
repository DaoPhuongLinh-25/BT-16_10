using WinFormsApp11;
using System;
using System.Windows.Forms;

namespace   WinFormsApp11
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}
