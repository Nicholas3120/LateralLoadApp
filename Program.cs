using System;
using System.Windows.Forms;

namespace LateralLoadApp
{
    static class Program
    {
        // The main entry point for the application
        [STAThread]
        static void Main()
        {
            // Enable visual styles and set compatibility with text rendering
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Run the MainForm as the starting point of the application
            Application.Run(new MainForm());
        }
    }
}
