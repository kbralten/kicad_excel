using System;
using System.Windows;

namespace KiCadExcelBridge
{
    public static class Program
    {
        [STAThread]
        public static int Main(string[] args)
        {
            try
            {
                var app = new App();
                app.InitializeComponent();
                app.Run();
                return 0;
            }
            catch (Exception ex)
            {
                try
                {
                    var log = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "startup-exception.txt");
                    System.IO.File.WriteAllText(log, ex.ToString());
                }
                catch { }
                Console.Error.WriteLine(ex.ToString());
                return 1;
            }
        }
    }
}
