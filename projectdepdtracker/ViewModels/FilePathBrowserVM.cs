using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Xml.Linq;
using System.Collections.ObjectModel;
using System.Xml;
using System.Data;
using System.Xml.XPath;
using ProjectDependencyTracker.Models;
using System.IO;
using System.ComponentModel;
using System.Threading;
using System.Collections.Specialized;
using System.Windows.Controls;
using System.Windows.Forms;
using DataGrid = System.Windows.Controls.DataGrid;
using System.Reflection;


namespace ProjectDependencyTracker.ViewModels
{
    public class FilePathBrowserVM : INotifyPropertyChanged
    {
        #region Properties
        private readonly string ItemGroup = "ItemGroup";
        private readonly string Project = "Project";
        private readonly string Include = "Include";

        private List<string> fileList;

        private string _drivePath;
        public string DrivePath
        {
            get => _drivePath;
            set
            {
                _drivePath = value;
                NotifyPropertyChanged(nameof(DrivePath));
            }
        }

        private string _projectName;
        public string ProjectName
        {
            get => _projectName;
            set
            {
                _projectName = value;
                NotifyPropertyChanged(nameof(ProjectName));
            }
        }

        protected IDataAccessLayer DAL { get; set; }

        public bool IsGetDependenciesEnabled { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        public void NotifyPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        public ICommand BrowseCommand { get; set; }

        public ICommand ExportCommand{ get; set; }

        public ICommand GetDependenciesCommand { get; set; }

        public static ObservableCollection<DependencyInfoModel> DependentsInfo { get; set; } = new ObservableCollection<DependencyInfoModel>();

        private bool CanExportCommandExecute(object obj)
        {
            return DependentsInfo?.Count > 0;
        }

        public void ExportCommandExecute(object obj)
        {
            try
            {
                using (SaveFileDialog dialog = new SaveFileDialog())
                {
                    if (dialog != null)
                    {
                        dialog.ShowDialog();
                        DataGrid dataGrid = obj as DataGrid;
                        var fileName = dialog.FileName + ".xlsx";
                        ExportProjDependencies(fileName, dataGrid);
                    }
                }
            }

            catch (Exception)
            {
            }
        }

        public static DataTable CreateTable<T>()
        {
            Type entityType = typeof(T);
            DataTable table = new DataTable(entityType.Name);
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(entityType);

            foreach (PropertyDescriptor prop in properties)
            {
                table.Columns.Add(prop.Name, prop.PropertyType);

                table.Rows.Add(DependentsInfo.Where(x => x.AdditionalIncludeDirectories == prop.Name));
                var data = DependentsInfo.Where(x => x.Contains(prop.Name));
                table.Rows.Add(DependentsInfo.Where(x => x.Contains(prop.Name)));
            }
            return table;
        }

        public void ExportProjDependencies(string filePath, DataGrid dg)
        {
            var table = CreateTable<DependencyInfoModel>();
            var list = new List<DependencyInfoModel>(dg.ItemsSource as IEnumerable<DependencyInfoModel>);
            var dataTable = ToDataTable(list);
            if (dataTable != null)
            {
                DataSet ds = new DataSet();
                ds.Tables.Add(dataTable);
                DAL.Write(filePath, ds);
            }
        }

        public static DataTable ToDataTable<T>(List<T> items)
        {            
            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
                
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        public void BrowserCommandExecute(object obj)
        {
            try
            {
                using (FolderBrowserDialog dialog = new FolderBrowserDialog())
                {
                    if (dialog != null)
                    {
                        dialog.SelectedPath = DrivePath;
                        dialog.ShowDialog();
                        DrivePath = dialog.SelectedPath;
                        if(!string.IsNullOrEmpty(DrivePath))
                        IsGetDependenciesEnabled = true;
                        NotifyPropertyChanged(nameof(IsGetDependenciesEnabled));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

        }

        //For getting the project files from a selected folder
        public void GetDependenciesCommandExecute(object obj)
        {
            if (!string.IsNullOrEmpty(DrivePath))
            {
                DependentsInfo.Clear();
                fileList = Directory.GetFiles(DrivePath, "*.vcxproj", SearchOption.AllDirectories).ToList();
                fileList.AddRange(Directory.GetFiles(DrivePath, "*.csproj", SearchOption.AllDirectories));
                foreach (var file in fileList)
                {
                    XmlDocument doc = new XmlDocument();
                    GetProjDependencies(file);                 
                }

                var all = DependentsInfo.FirstOrDefault(x => x.AdditionalIncludeDirectories == x.AdditionalIncludeDirectories);
               
            }
            NotifyPropertyChanged(nameof(ExportCommand));
        }
                   
        private void GetProjDependencies(string fileName)
        {
            XNamespace msbuild = "http://schemas.microsoft.com/developer/msbuild/2003";
            XmlDocument doc = new XmlDocument();
            doc.Load(fileName);
            if (doc.DocumentElement != null)
            {
                XmlElement root = doc.DocumentElement;
                List<XmlNodeList> nodeList = new List<XmlNodeList>
            {
                root.GetElementsByTagName("ClCompile"),
                root.GetElementsByTagName("Link"),
                root.GetElementsByTagName("ClInclude ")
            };
                DependencyInfoModel objDepInfo = new DependencyInfoModel
                {
                    ProjectNo = (fileList.IndexOf(fileName) + 1).ToString(),
                    ProjectPath = fileName,
                    DeliverableName = Path.GetFileNameWithoutExtension(fileName)
                };

                foreach (XmlNodeList titleNode in nodeList)
                {
                    XmlNode xmlNode = titleNode[0];

                    if (xmlNode != null)
                    {
                        XmlNodeList titleList = xmlNode.ChildNodes;

                        foreach (XmlNode node in titleList)
                        {
                            if (node.Name.Equals("OutputFile"))
                            {
                                objDepInfo.DeliverableName = string.Join("", node.InnerText.Skip(9));
                            }

                            if (node.Name.Equals("AdditionalDependencies"))
                            {
                                string splitNodes = node.InnerText.Split('%').First();
                                objDepInfo.LinkerDependency = string.Join<string>(Environment.NewLine, splitNodes.Split(';').Distinct());
                            }

                            if (node.Name.Equals("AdditionalIncludeDirectories"))
                            {
                                string splitNodes = node.InnerText.Split('%').First();
                                objDepInfo.AdditionalIncludeDirectories = string.Join<string>(Environment.NewLine, splitNodes.Split(';').Distinct());
                            }
                        }
                    }

                    if (fileName.EndsWith(".csproj"))
                    {
                        try
                        {
                            XDocument projElements = XDocument.Load(fileName);
                            List<string> itemGroupReferences = projElements
                                   .Element(msbuild + Project)
                                   .Elements(msbuild + ItemGroup)
                                   .Elements(msbuild + "Reference")
                                   .Elements(msbuild + "HintPath")
                                   .Select(refElem => refElem.Value).ToList();

                            itemGroupReferences.AddRange(projElements
                                 .Element(msbuild + Project)
                                 .Elements(msbuild + ItemGroup)
                                 .Elements(msbuild + "None")
                                 .Attributes(Include)
                                 .Select(refElem => refElem.Value).Where(x => x.StartsWith(".")));

                            itemGroupReferences.AddRange(projElements
                               .Element(msbuild + Project)
                               .Elements(msbuild + ItemGroup)
                               .Elements(msbuild + "ProjectReference")
                               .Attributes(Include)
                               .Select(refElem => refElem.Value).Where(x => x.StartsWith(".")));

                            itemGroupReferences.AddRange(projElements
                           .Element(msbuild + Project)
                           .Elements(msbuild + ItemGroup)
                           .Elements(msbuild + "Compile")
                           .Attributes(Include)
                           .Select(refElem => refElem.Value).Where(x => x.StartsWith(".")));

                            itemGroupReferences.AddRange(projElements
                            .Element(msbuild + Project)
                            .Elements(msbuild + ItemGroup)
                            .Elements(msbuild + "EmbeddedResource")
                            .Attributes(Include)
                            .Select(refElem => refElem.Value).Where(x => x.StartsWith(".")));

                            var referencedDlls = itemGroupReferences.Where(x => x.EndsWith(".dll") || x.EndsWith(".csproj"));
                            var referencedFiles = itemGroupReferences.Except(referencedDlls);
                            objDepInfo.ReferenceComponents = string.Join<string>(Environment.NewLine, referencedDlls);
                            objDepInfo.ReferenceFiles = string.Join<string>(Environment.NewLine, referencedFiles);
                            objDepInfo.ProjectType = "C#";
                        }
                        catch(Exception ex)
                        {
                            Console.WriteLine(ex);
                        }
                    }
                    else
                        objDepInfo.ProjectType = "c++";

                }
                DependentsInfo.Add(objDepInfo);
            }
        }

        public void ExporttoExcel(string strFilePath)
        {
            try
            {
                var fileName = strFilePath + "\\DependencyTrackerReport" + DateTime.Now.ToString("dd-MM-yy hh-mm") + ".xlsx";
                var table = CreateTable<DependencyInfoModel>();
                var list = new List<DependencyInfoModel>(DependentsInfo as IEnumerable<DependencyInfoModel>);
                var dataTable = ToDataTable(list);
                if (dataTable != null)
                {
                    DataSet ds = new DataSet();
                    ds.Tables.Add(dataTable);
                    DAL.Write(fileName, ds);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Environment.Exit(-1);
            }
        }

        public void GetProjectDependencies()
        {
            if (!string.IsNullOrEmpty(DrivePath))
            {
                DependentsInfo.Clear();
                fileList = Directory.GetFiles(DrivePath, "*.vcxproj", SearchOption.AllDirectories).ToList();
                fileList.AddRange(Directory.GetFiles(DrivePath, "*.csproj", SearchOption.AllDirectories));
                foreach (var file in fileList)
                {
                    XmlDocument doc = new XmlDocument();
                    GetProjDependencies(file);
                }
                var all = DependentsInfo.FirstOrDefault(x => x.AdditionalIncludeDirectories == x.AdditionalIncludeDirectories);

            }
        }

        public FilePathBrowserVM()
        {
            BrowseCommand = new RelayCommand(BrowserCommandExecute);
            GetDependenciesCommand = new RelayCommand(GetDependenciesCommandExecute);
            ExportCommand = new RelayCommand(ExportCommandExecute, CanExportCommandExecute);
            DependentsInfo = new ObservableCollection<DependencyInfoModel>();
            DAL = new DataAccessLayer();
        }

        public class RelayCommand : ICommand
        {
            private Action<object> execute;

            private Predicate<object> canExecute;

            private event EventHandler CanExecuteChangedInternal;

            public RelayCommand(Action<object> execute)
                : this(execute, DefaultCanExecute)
            {
            }

            public RelayCommand(Action<object> execute, Predicate<object> canExecute)
            {
                if (execute == null)
                {
                    throw new ArgumentNullException("execute");
                }

                if (canExecute == null)
                {
                    throw new ArgumentNullException("canExecute");
                }

                this.execute = execute;
                this.canExecute = canExecute;
            }

            public event EventHandler CanExecuteChanged
            {
                add
                {
                    CommandManager.RequerySuggested += value;
                    this.CanExecuteChangedInternal += value;
                }

                remove
                {
                    CommandManager.RequerySuggested -= value;
                    this.CanExecuteChangedInternal -= value;
                }
            }

            public bool CanExecute(object parameter)
            {                
                return this.canExecute != null && this.canExecute(parameter);
            }

            public void Execute(object parameter)
            {
                this.execute(parameter);
            }

            public void OnCanExecuteChanged()
            {
                EventHandler handler = this.CanExecuteChangedInternal;
                if (handler != null)
                {
                    //DispatcherHelper.BeginInvokeOnUIThread(() => handler.Invoke(this, EventArgs.Empty));
                    handler.Invoke(this, EventArgs.Empty);
                }
            }

            public void Destroy()
            {
                this.canExecute = _ => false;
                this.execute = _ => { return; };
            }

            private static bool DefaultCanExecute(object parameter)
            {
                return true;
            }

            

        }
    }
}
