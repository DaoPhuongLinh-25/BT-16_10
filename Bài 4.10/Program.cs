using Bài_16_10__1_;
using System;
using System.Windows.Forms;

namespace Bài_16_10__1_
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}