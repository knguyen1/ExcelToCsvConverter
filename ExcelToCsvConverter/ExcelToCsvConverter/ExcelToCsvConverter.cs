using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Data.OleDb;
using System.Data;

namespace ExcelToCsvConverter
{
    class ExcelToCsvConverter
    {
        //member variables
        private DirectoryInfo _filePath;
        private IEnumerable<FileInfo> _excelFiles;
        //private string _connectionString;

        public ExcelToCsvConverter(string filePath)
        {
            string path = Path.GetFullPath(filePath);
            _filePath = new DirectoryInfo(path);

            _excelFiles = _filePath.GetFilesByExtensions(".xlsx",".xls");
        }

        /// <summary>
        /// Converts the sheet at position number to csv.
        /// </summary>
        /// <param name="worksheetNumber">The position of the sheet, 1-based.</param>
        public void ConvertSheet(int worksheetNumber)
        {
            if (!_excelFiles.Any())
                throw new ArgumentNullException("There are no Excel files in this directory.");

            foreach (var file in _excelFiles)
            {
                string connectionString;

                if(Path.GetExtension(file.FullName) == ".xls")
                    connectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;IMEX=1;HDR=NO\"", file.FullName);
                else
                    connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=NO\"", file.FullName);

                using (var connection = new OleDbConnection(connectionString))
                using (DataTable table = new DataTable())
                {
                    //get schema, then data
                    string worksheet;
                    try
                    {
                        if (connection.State != ConnectionState.Open)
                            connection.Open();

                        using (DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null))
                        {
                            if (schemaTable.Rows.Count < worksheetNumber)
                                throw new ArgumentException("The worksheet number provided cannot be found in the spreadsheet.");

                            worksheet = schemaTable.Rows[worksheetNumber - 1]["table_name"].ToString().Replace("'", "");
                            string sql = String.Format("SELECT * FROM [{0}]", worksheet);

                            var dataAdapter = new OleDbDataAdapter(sql, connection);
                            dataAdapter.Fill(table);
                        }
                    }
                    catch (Exception exc)
                    {
                        throw exc;
                    }

                    //check if the first row is empty
                    DataRow firstRow = table.Rows[0];
                    if (IsRowEmpty(firstRow))
                        table.Rows.Remove(firstRow);

                    //finally, write to table
                    string fileName = Path.GetFileNameWithoutExtension(file.FullName);
                    WriteToCsv(table, fileName + "-" + worksheet.Replace("$", String.Empty));
                }
            }
        }

        /// <summary>
        /// Converts the sheet with a given sheet name.
        /// </summary>
        /// <param name="worksheetNumber">The string indicating the sheet name.</param>
        public void ConvertSheet(string sheetName)
        {
            if (!_excelFiles.Any())
                throw new ArgumentNullException("There are no Excel files in this directory.");

            foreach (var file in _excelFiles)
            {
                string connectionString;

                if (Path.GetExtension(file.FullName) == ".xls")
                    connectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;IMEX=1;HDR=NO\"", file.FullName);
                else
                    connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=NO\"", file.FullName);

                using (var connection = new OleDbConnection(connectionString))
                using (DataTable table = new DataTable())
                {
                    //get schema, then data
                    try
                    {
                        if (connection.State != ConnectionState.Open)
                            connection.Open();

                        string sql = String.Format("SELECT * FROM [{0}]", sheetName);

                        var dataAdapter = new OleDbDataAdapter(sql, connection);
                        dataAdapter.Fill(table);
                    }
                    catch (Exception exc)
                    {
                        throw exc;
                    }

                    //check if the first row is empty
                    DataRow firstRow = table.Rows[0];
                    if (IsRowEmpty(firstRow))
                        table.Rows.Remove(firstRow);

                    //finally, write to table
                    string fileName = Path.GetFileNameWithoutExtension(file.FullName);
                    WriteToCsv(table, fileName + "-" + sheetName);
                }
            }
        }

        /// <summary>
        /// Converts all sheets (be careful!) in the directory.
        /// </summary>
        public void ConvertAllSheets()
        {
            if (!_excelFiles.Any())
                throw new ArgumentNullException("There are no Excel files in this directory.");

            foreach (var file in _excelFiles)
            {
                string connectionString;

                if (Path.GetExtension(file.FullName) == ".xls")
                    connectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;IMEX=1;HDR=NO\"", file.FullName);
                else
                    connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=NO\"", file.FullName);

                using (var connection = new OleDbConnection(connectionString))
                using (DataTable table = new DataTable())
                {
                    //get schema, then data
                    try
                    {
                        if (connection.State != ConnectionState.Open)
                            connection.Open();

                        using (DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null))
                        {
                            foreach (var schemaRow in schemaTable.AsEnumerable().Cast<DataRow>().Where(row => row["TABLE_NAME"].ToString().EndsWith("$")))
                            {
                                //Console.WriteLine(schemaRow.ItemArray[2].ToString().Replace("'", ""));
                                ConvertSheet(schemaRow.ItemArray[2].ToString().Replace("'", ""));
                            }
                        }
                    }
                    catch (Exception exc)
                    {
                        throw exc;
                    }
                }
            }
        }

        /// <summary>
        /// Writers the DataTable object to a csv file.
        /// </summary>
        /// <param name="table">The DataTable object.</param>
        /// <param name="file">The FileInfo object of the input.</param>
        private void WriteToCsv(DataTable table, string fileName)
        {
            //string fileName = Path.GetFileNameWithoutExtension(file.FullName);
            string output = _filePath + "\\" + fileName + "-csv.csv";

            using (var writer = new StreamWriter(output))
            {
                bool headerRow = true;
                foreach (DataRow row in table.Rows)
                {
                    bool firstCol = true;
                    foreach (DataColumn col in table.Columns)
                    {
                        if (!firstCol)
                            writer.Write("\t");
                        else
                            firstCol = false;

                        string data = row[col.ColumnName].ToString().Trim();

                        //if it's a header row, it will contain escape characters you need to trim
                        if (headerRow)
                        {
                            char[] escapeChars = new[] { '\n', '\a', '\r' };
                            string clean = new string(data.Where(c => !escapeChars.Contains(c)).ToArray());
                            writer.Write(clean);
                        }
                        else
                        {
                            writer.Write(data);
                        }
                    }

                    headerRow = false;

                    //check if it's the last row, if not then write a new line
                    if(row != table.Rows[table.Rows.Count - 1])
                        writer.WriteLine();
                }
            }
        }

        /// <summary>
        /// Tells whether or not the DataRow provided is empty.
        /// </summary>
        /// <param name="dr">The DataRow parameter</param>
        /// <returns>True/False if it's empty.</returns>
        private bool IsRowEmpty(DataRow dr)
        {
            if (dr == null)
                return true;
            else
            {
                foreach (var value in dr.ItemArray)
                {
                    if (value.ToString() != String.Empty)
                        return false;
                }

                return true;
            }
        }
    }
}
