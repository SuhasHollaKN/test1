using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ProjectDependencyTracker.ViewModels
{
    class DataAccessLayer : IDataAccessLayer
    {
        private IDataAccessLayer DAL = null;

        #region DataAccessLayer Constructors

        /// <summary>
        /// Constrcutor for the DataAccessLayer Class
        /// </summary>
        public DataAccessLayer()
        {

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataAccessLayer"/> class.
        /// </summary>
        private void InitializeDataAccessLayer(string fileName)
        {
            DAL = null;
            string strFilePath = fileName.ToUpper();

            if (strFilePath.EndsWith(".XLSX"))
            {
                DAL = new OpenXMLExcelParser();
            }
            else
            {
                //throw new BusinessLogicException(Resources.FileFormatErr);
            }
        }

        #endregion

        #region DataAccessLayer Methods

        /// <summary>
        /// Reads all the worksheets from the specified file
        /// </summary>
        public DataSet Read(string fileName)
        {
            InitializeDataAccessLayer(fileName);
            return DAL.Read(fileName);
        }

        /// <summary>
        /// Read the data from the file specified
        /// </summary>
        public DataTable Read(string fileName, ArrayList arrSheets)
        {
            InitializeDataAccessLayer(fileName);
            return DAL.Read(fileName, arrSheets);
        }

        /// <summary>
        /// Writes to sheet per cell
        /// </summary>
        public void Write(string fileName, DataSet dsTables)
        {
            InitializeDataAccessLayer(fileName);
            DAL.Write(fileName, dsTables);
        }

        /// <summary>
        /// Writes to sheet per cell
        /// </summary>
        public void Write(string fileName, DataTable dtSelectedRows)
        {
            InitializeDataAccessLayer(fileName);
            DAL.Write(fileName, dtSelectedRows);
        }        

        #endregion

        public class OpenXMLExcelParser : IDataAccessLayer
        {

            public DataSet Read(string fileName)
            {
                return ExtractExcelSheetValuesToDataSet(fileName);

                // throw new NotImplementedException();
            }

            public DataTable Read(string fileName, ArrayList arrSheets)
            {
                throw new NotImplementedException();
            }

            public void Write(string fileName, DataSet dsTables)
            {
                ExportDSToExcel(dsTables, fileName);
                //throw new NotImplementedException();
            }

            public void Write(string fileName, DataTable dtSelectedRows)
            {

                throw new NotImplementedException();
            }



            private DataSet ExtractExcelSheetValuesToDataSet(string xlsxFilePath)
            {

                DataSet ds = new DataSet();

                using (var spreadSheetDocument = SpreadsheetDocument.Open(xlsxFilePath, true))
                {
                    WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;

                    int sheetIndex = 0;
                    foreach (Sheet s in workbookPart.Workbook.Descendants<Sheet>())
                    {
                        WorksheetPart worksheetpart = workbookPart.GetPartById(s.Id) as WorksheetPart;

                        if (worksheetpart == null)
                        {
                            continue;
                        }


                        DataTable dt = new DataTable();
                        Worksheet worksheet = worksheetpart.Worksheet;
                        dt.TableName = s.Name;

                        IEnumerable<Row> rows = worksheetpart.Worksheet.Descendants<Row>();
                        int noofrows = rows.Count();

                        //If there is an empty sheet in the workbook just add an empty sheet 
                        //for example notes sheet in EDB_Template doesn't contain any data so add new notes sheet with empty rows to dateset
                        if (rows.Count() == 0)
                        {
                            ds.Tables.Add(dt);
                            continue;
                        }

                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in rows.ElementAt(0))
                        {
                            string value = GetCellValue(spreadSheetDocument, cell);
                            if (dt.Columns.Contains(value))
                            {
                                throw new Exception("Failed to read workbook: " + s.Name + " contain the duplicate column " + value + " Please remove all duplicate column from the workbook. ");
                            }
                            else
                            {
                                dt.Columns.Add(value);
                            }
                        }

                        int temprow = 0;
                        foreach (Row row in rows) //this will also include your header row...
                        {
                            DataRow tempRow = dt.NewRow();
                            //for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                            {
                                int columnIndex = 0;
                                foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in row.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                                {
                                    // Gets the column index of the cell with data
                                    int cellColumnIndex = (int)GetColumnIndexFromName(GetColumnName(cell.CellReference));
                                    // if there is any empty cells in the row Open xml will not read this data  
                                    // so get the index for the next value and add blank data in between
                                    if (columnIndex < cellColumnIndex)
                                    {
                                        do
                                        {

                                            if (columnIndex >= dt.Columns.Count)
                                            {
                                                //If any data (might be courpted ) in Excel after last column, skip the addition of this data
                                                break;
                                            }
                                            tempRow[columnIndex] = string.Empty;//Insert blank data here;
                                            columnIndex++;
                                        }
                                        while (columnIndex < cellColumnIndex);
                                    }

                                    if (columnIndex <= dt.Columns.Count)
                                    {
                                        //Adding the data till last param only, if there is no param name it is not requird to process
                                        tempRow[columnIndex - 1] = GetCellValue(spreadSheetDocument, cell);
                                        columnIndex++;
                                    }
                                }

                            }

                            //below code allows only one empty row between all the available rows, if there are consecutive empty rows, than those will not be added back to sheet. 
                            bool isEmpty = IsDataRowEmpty(tempRow);
                            if (!isEmpty)
                            {
                                //adding the first empty row
                                dt.Rows.Add(tempRow);
                            }
                            else
                            {
                                //skip the consecutive empty rows
                                if (temprow != dt.Rows.Count)
                                {
                                    dt.Rows.Add(tempRow);
                                    temprow = dt.Rows.Count;
                                }

                            }
                        }
                        sheetIndex++;

                        if (dt.Rows.Count != 0)
                        {
                            dt.Rows.RemoveAt(0); //first row which has column names is repeated, hence removing first row here.
                            ds.Tables.Add(dt);
                        }
                    }
                    spreadSheetDocument.Close();
                }

                return ds;
            }


            private string GetColumnName(string cellReference)
            {
                if (cellReference == null)
                    return null;

                // Create a regular expression to match the column name portion of the cell name.
                Regex regex = new Regex("[A-Za-z]+");
                Match match = regex.Match(cellReference);

                return match.Value;
            }

            // This Function will return the Column number for the given column name
            // for A1 column it will return 1 and B1 it will be 2 similarly for AA1 it will be 27
            //if user is processing with zero based index he use columnIndex-1 for adding the values
            private int GetColumnIndexFromName(string columnName)
            {
                if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

                columnName = columnName.ToUpperInvariant();

                int columnIndex = 0;

                for (int i = 0; i < columnName.Length; i++)
                {

                    columnIndex *= 26;
                    columnIndex += (columnName[i] - 'A' + 1); // Added +1 here for getting correct indexes from AA onwards.
                }
                return columnIndex;
            }

            private bool IsDataRowEmpty(DataRow dr)
            {
                if (dr == null)
                {
                    return true;
                }
                else
                {
                    foreach (var value in dr.ItemArray)
                    {
                        if (value != null && value != DBNull.Value)
                        {
                            if (!String.IsNullOrEmpty((String)value))
                            {
                                return false;
                            }
                        }
                    }
                    return true;
                }
            }
            private string GetCellValue(SpreadsheetDocument document, Cell cell)
            {
                SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                string value = string.Empty;
                if (cell.CellValue != null)
                {
                    value = cell.CellValue.InnerXml;

                    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        value = stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                    }
                    else
                    {
                        value = cell.CellValue.InnerXml;
                    }
                }

                return value;

            }

            public static string GetStandardExcelColumnName(int columnNumberOneBased)
            {
                int baseValue = Convert.ToInt32('A');
                int columnNumberZeroBased = columnNumberOneBased - 1;

                string ret = "";

                if (columnNumberOneBased > 26)
                {
                    ret = GetStandardExcelColumnName(columnNumberZeroBased / 26);
                }

                return ret + Convert.ToChar(baseValue + (columnNumberZeroBased % 26));
            }

            private void ExportDSToExcel(DataSet ds, string destination)
            {

                try
                {

                    if (File.Exists(destination))
                    {
                        File.Delete(destination);
                    }


                    using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                    {
                        var workbookPart = workbook.AddWorkbookPart();
                        workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                        workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                        uint sheetId = 1;

                        foreach (DataTable table in ds.Tables)
                        {
                            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                            var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                            sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                            DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                            if (sheets.Elements<Sheet>().Count() > 0)
                            {
                                sheetId =
                                    sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                            }

                            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                            sheets.Append(sheet);

                            DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                            int rowIndex = 1;
                            List<String> columns = new List<string>();
                            foreach (DataColumn column in table.Columns)
                            {
                                columns.Add(column.ColumnName);

                                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                                cell.CellReference = GetStandardExcelColumnName(rowIndex);
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                                headerRow.AppendChild(cell);
                                rowIndex++;
                            }

                            sheetData.AppendChild(headerRow);

                            foreach (DataRow dsrow in table.Rows)
                            {
                                int columnIndex = 1;
                                DocumentFormat.OpenXml.Spreadsheet.Row newRow = new Row();
                                foreach (String col in columns)
                                {
                                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                                    cell.CellReference = GetStandardExcelColumnName(columnIndex);
                                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                                    newRow.AppendChild(cell);
                                    columnIndex++;
                                }

                                sheetData.AppendChild(newRow);
                            }

                        }
                        workbook.Close();
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

        }
    }
}
