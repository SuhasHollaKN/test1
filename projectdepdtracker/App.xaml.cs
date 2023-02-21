using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using ProjectDependencyTracker.ViewModels;


namespace ProjectDependencyTracker
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public void Application_Startup(object sender, StartupEventArgs e)
        {
            if (e.Args.Length == 0)
            {
                MainWindow MainUI = new MainWindow();
                MainUI.Show();
            }
            else if (e.Args.Length == 2)
            {
                FilePathBrowserVM objFileBrowser = new FilePathBrowserVM();
                objFileBrowser.DrivePath = e.Args[0];

                objFileBrowser.GetProjectDependencies();
                //Requrest Excel Path
                objFileBrowser.ExporttoExcel(e.Args[1].Replace("\"",""));
                Environment.Exit(0);
            }
            else
            {
                Console.WriteLine("Invalid Arguments. Please run the utility as shown below \n" +
                    "ProjectDependencyTracker.exe 'Path of the Project(s)' 'Path to save Report' ");
                Environment.Exit(-1);
            }
        }
    }
}
