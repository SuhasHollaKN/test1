using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectDependencyTracker.Models
{
    public class DependencyInfoModel : IEnumerable<string>
    {
        public DependencyInfoModel()
        {
        }
        public string ProjectNo { get; set; }
        public string ProjectPath { get; set; }
        public string ProjectType { get; set; }
        public string DeliverableName { get; set; }
        public string LinkerDependency { get; set; }
        public string AdditionalIncludeDirectories { get; set; }
        public string ReferenceFiles { get; set; }
        public string ReferenceComponents { get; set; }

        private IEnumerable<string> Dependents()
        {
            yield return ProjectType;
            yield return ProjectPath;
            yield return DeliverableName;
            yield return LinkerDependency;
            yield return AdditionalIncludeDirectories;
            yield return ReferenceFiles;
            yield return ReferenceComponents;
            yield return ProjectNo;

        }

        public IEnumerator<string> GetEnumerator()
        {
            return Dependents().GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }

    
}
