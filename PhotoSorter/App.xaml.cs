using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace PhotoSorter
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        void App_Startup(object sender, StartupEventArgs e)
        {
            // Application is running
            // Process command line args
            bool RunImport = false;
            bool badArg = false;
            for (int i = 0; i != e.Args.Length; ++i)
            {
                if (e.Args[i].ToLower().Trim() == "/runimport" || e.Args[i].ToLower().Trim() == "-runimport")
                {
                    RunImport = true;
                }
                else
                {
                    badArg = true;
                }
            }

            if (!badArg)
            {
                if (RunImport)
                {
                    Importer imp = new Importer();
                    Console.WriteLine("Running Photo import...");
                    imp.RunImport();
                    Console.WriteLine("Photo Import complete.");
                    Application.Current.Shutdown();
                }
                else
                {
                    MainWindow mainWindow = new MainWindow();
                    mainWindow.Show();
                }
            }
            else
            {
                Console.WriteLine("Bad argument has been supplied.  To run the import include argument \"\\/RunImport\"");
                Application.Current.Shutdown();
            }
        }
    }
}
