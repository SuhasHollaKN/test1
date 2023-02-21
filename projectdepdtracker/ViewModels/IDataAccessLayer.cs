using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectDependencyTracker.ViewModels
{
    public interface IDataAccessLayer
    {
        /// <summary>
        /// Reads all the worksheets from the specified file
        /// </summary>
        DataSet Read(string fileName);

        /// <summary>
        /// Read the given list of worksheets from the file
        /// </summary>
        DataTable Read(string fileName, ArrayList arrSheets);

        /// <summary>
        /// Writes all the Tables in a dataset to the file
        /// </summary>
        void Write(string fileName, DataSet dsTables);

        /// <summary>
        /// Writes the table to the file
        /// </summary>
        void Write(string fileName, DataTable dtSelectedRows);

    }
}

