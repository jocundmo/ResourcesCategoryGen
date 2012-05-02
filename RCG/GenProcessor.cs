using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Collections.ObjectModel;

namespace RCG
{
    public class GenProcessor
    {
        #region Constants

        #endregion

        #region Constructor

        public GenProcessor()
        {
            // Nothing to do.
        }

        #endregion

        #region Private members

        private ConfigDoc _config = new ConfigDoc();
        private DataSet _metadataSet = new DataSet();
        private DataSet _excelSet = new DataSet();
        private Excel.Application _excel;

        #endregion

        #region Properties

        public ConfigDoc Config
        {
            get { return _config; }
        }

        public SheetConfig CurrentSheetConfig { get; private set; }

        public ColumnConfig CurrentColumnConfig { get; private set; }

        public dynamic CurrentActiveExcelSheet { get; private set; }

        public DataSet MetadataSet { get { return _metadataSet; } }

        public DataSet ExcelSet { get { return _excelSet; } }

        #endregion

        #region Helper Methods

        private SheetConfig FindSheetConfig(ConfigDoc config, string sheetName)
        {
            foreach (SheetConfig sheetConfig in config.Sheets)
            {
                if (sheetConfig.Name == sheetName)
                {
                    CurrentSheetConfig = sheetConfig;
                    return sheetConfig;
                }
            }
            throw new Exception(string.Format("Sheet config {0} not recognized...", sheetName));
        }

        private ColumnConfig FindColumnConfig(ConfigDoc config, string sheetName, string columnName)
        {
            SheetConfig sheetConfig = FindSheetConfig(config, sheetName);

            foreach (ColumnConfig columnConfig in sheetConfig.Columns)
            {
                if (columnConfig.Name == columnName)
                {
                    CurrentColumnConfig = columnConfig;
                    return columnConfig;
                }
            }
            throw new Exception(string.Format("Column config {0} of Sheet {1} is not recognized...", columnName, sheetName));
        }

        #endregion

        #region Instance Methods

        /// <summary>
        /// Reads the configuration within "Mappings.xml"
        /// </summary>
        /// <param name="configFileName"></param>
        public void ReadConfiguration(string configFileName)
        {
            _config.Read(configFileName);
            _config.Validate();
        }
        public void ReadConfiguration()
        {
            ReadConfiguration("Mappings.xml");
        }

        /// <summary>
        /// Reads the data which resideds in your harddisk and generated metadata based on it.
        /// </summary>
        public void GenerateMetadata()
        {
            foreach (SheetConfig sheetConfig in _config.Sheets)
            {
                if (!sheetConfig.Enabled)
                    continue;

                CurrentSheetConfig = sheetConfig;

                DataTable dt = new DataTable(sheetConfig.Name);

                // Generate the original Columns.
                DataColumn dcPath = new DataColumn(Constants.COLUMN_Path);
                dt.Columns.Add(dcPath);
                DataColumn dcLastModified = new DataColumn(Constants.COLUMN_LastModified);
                dt.Columns.Add(dcLastModified);
                DataColumn dcSize = new DataColumn(Constants.COLUMN_Size);
                dt.Columns.Add(dcSize);
                DataColumn dcAttributes = new DataColumn(Constants.COLUMN_Attributes);
                dt.Columns.Add(dcAttributes);
                DataColumn dcFileCount = new DataColumn(Constants.COLUMN_FileCount);
                dt.Columns.Add(dcFileCount);
                DataColumn dcSubFolderCount = new DataColumn(Constants.COLUMN_SubFolderCount);
                dt.Columns.Add(dcSubFolderCount);
                DataColumn dcFromFile = new DataColumn(Constants.COLUMN_FromFile);
                dt.Columns.Add(dcFromFile);
                DataColumn dcFilesType = new DataColumn(Constants.COLUMN_FilesType);
                dt.Columns.Add(dcFilesType);
                DataColumn dcFilterFlag = new DataColumn(Constants.COLUMN_FilterFlag, typeof(bool));
                dt.Columns.Add(dcFilterFlag);
                DataColumn dcRowMode = new DataColumn(Constants.COLUMN_RowMode);
                dt.Columns.Add(dcRowMode);
                DataColumn dcFormatString = new DataColumn(Constants.COLUMN_Formatter);
                dt.Columns.Add(dcFormatString);
                DataColumn dcTag = new DataColumn(Constants.COLUMN_Tag);
                dt.Columns.Add(dcTag);

                // Read metadata.
                foreach (LocationConfig locationConfig in sheetConfig.Locations)
                {
                    if (!locationConfig.Enabled)
                        continue;
                    try
                    {
                        ReadMetadata(locationConfig, dt);
                    }
                    catch (IOException ex)
                    {
                        if (OnHandlableException != null)
                            OnHandlableException(this, new HandlableExceptionEventArgs(ex, string.Empty));
                    }
                }

                _metadataSet.Tables.Add(dt);
            }
        }

        private void ReadMetadata(LocationConfig locationConfig, DataTable dt)
        {
            const string FROM_FILE_NAME = "\\metadata.txt";

            string sPath = locationConfig.Path;
            // Supports use VolumeLabel directly instead of physical driver name.
            LocationType lt = Utility.GetLocationType(sPath);
            if (lt == LocationType.VolumeLabel)
            {
                sPath = Utility.ConvertVolumeLabelPathToPhysicalPath(locationConfig.Path);
            }
            string[] dirList = null;

            if (string.IsNullOrEmpty(sPath) || !Directory.Exists(sPath))
                return;

            if (locationConfig.IncludeFolder)
            {
                dirList = Directory.GetDirectories(sPath);
            }
            else
            {
                DirectoryInfo di = new DirectoryInfo(sPath);

                Collection<string> sCollection = new Collection<string>();
                foreach (FileTypeConfig s in locationConfig.IncludeFileTypes)
                {
                    FileInfo[] fiArray = di.GetFiles(s.FileTypeExtension, s.SearchOption);
                    foreach (FileInfo fi in fiArray)
                    {
                        sCollection.Add(fi.FullName);
                    }
                }
                dirList = sCollection.ToArray();
            }
            foreach (string dir in dirList)
            {
                DirectoryInfo d = new DirectoryInfo(dir);
                DataRow dr = dt.NewRow();

                dr[Constants.COLUMN_Path] = dir;
                dr[Constants.COLUMN_LastModified] = d.LastWriteTime.ToString();
                dr[Constants.COLUMN_Attributes] = d.Attributes.ToString();

                long size = 0;
                int fileCount = 0;
                int subFolderCount = 0;
                string filesType = string.Empty;
                Utility.CalFolderSize(ref size, ref fileCount, ref subFolderCount, ref filesType, d);

                dr[Constants.COLUMN_FilesType] = filesType;
                dr[Constants.COLUMN_Size] = size.ToString();
                dr[Constants.COLUMN_FileCount] = fileCount.ToString();
                dr[Constants.COLUMN_SubFolderCount] = subFolderCount.ToString();

                string fileContent = string.Empty;
                if (File.Exists(dir + FROM_FILE_NAME))
                {
                    using (StreamReader sr = new StreamReader(dir + FROM_FILE_NAME, Encoding.Default))
                    {
                        fileContent = sr.ReadToEnd();
                    }
                }
                dr[Constants.COLUMN_FromFile] = fileContent;

                if (OnReadingMetadata != null)
                    OnReadingMetadata(this, new DataRowEventArgs(dr, dir));

                dt.Rows.Add(dr);
            }
        }

        public void OutputTemporaryFiles(string filename)
        {
            _metadataSet.WriteXmlSchema(filename + ".xsd");
            _metadataSet.WriteXml(filename);
        }
        /// <summary>
        /// Reads the original existing excel to determine what already exists.
        /// </summary>
        public void ReadPreviousMetadata(string filename)
        {
            //_excel = new Excel.Application();
            //// Inits the active sheet.
            //if (!string.IsNullOrEmpty(_config.BaselinePath.Trim()) &&
            //    File.Exists(_config.BaselinePath.Trim()))
            //{
            //    _excel.Application.Workbooks.Open(_config.BaselinePath.Trim());
            //}
            //else if (!string.IsNullOrEmpty(_config.OutputPath.Trim()) &&
            //    File.Exists(_config.OutputPath.Trim()))
            //{
            //    _excel.Application.Workbooks.Open(_config.OutputPath.Trim());
            //}
            //else
            //{
            //    if (!string.IsNullOrEmpty(_config.TemplatePath.Trim()) &&
            //        File.Exists(_config.TemplatePath.Trim()))
            //        _excel.Application.Workbooks.Open(_config.TemplatePath.Trim());
            //    else
            //        _excel.Application.Workbooks.Add(true);
            //}

            //foreach (SheetConfig sheetConfig in _config.Sheets)
            //{
            //    if (!sheetConfig.Enabled)
            //        continue;

            //    CurrentSheetConfig = sheetConfig;
            //    CurrentActiveExcelSheet = ExcelOperationWrapper.FindExcelActiveSheet(_excel, sheetConfig.Name);

            //    DataTable excelTable = new DataTable(sheetConfig.Name);
            //    // Excel table columns must be generated by ColumnConfig, cant be Excel.
            //    foreach (ColumnConfig columnConfig in sheetConfig.Columns)
            //    {
            //        if (!columnConfig.Enabled)
            //            continue;
            //        excelTable.Columns.Add(new DataColumn(columnConfig.Name));
            //    }

            //    // Gets actual Excel col count
            //    int actualExcelColCount = 0;
            //    for (int colIndex = 1; colIndex <= CurrentSheetConfig.Columns.Count + 1; colIndex++)
            //    {
            //        if (CurrentActiveExcelSheet.Cells[Constants.HEADER_ROW_INDEX, colIndex].Value == null)
            //        {
            //            actualExcelColCount = colIndex - 1;
            //            break;
            //        }
            //    }

            //    // Gets primary column index from Excel.
            //    int primaryExcelColumnIndex = 0;
            //    if (actualExcelColCount > 0)
            //    {
            //        foreach (ColumnConfig columnConfig in CurrentSheetConfig.Columns)
            //        {
            //            if (!columnConfig.Primary)
            //                continue;

            //            for (int excelColIndex = 1; excelColIndex <= actualExcelColCount; excelColIndex++)
            //            {
            //                if (CurrentActiveExcelSheet.Cells[Constants.HEADER_ROW_INDEX, excelColIndex].Value == columnConfig.DisplayName)
            //                {
            //                    primaryExcelColumnIndex = excelColIndex;
            //                    break;
            //                }
            //            }
            //        }
            //    }

            //    // Gets actual Excel row count without header.
            //    int actualExcelRowCountWithoutHeader = 0;
            //    if (primaryExcelColumnIndex > 0)
            //    {
            //        int assumedMaxRowsCount = CurrentSheetConfig.MaxRowCount;

            //        for (int rowIndex = assumedMaxRowsCount; rowIndex > Constants.HEADER_ROW_INDEX; rowIndex--)
            //        {
            //            if (CurrentActiveExcelSheet.Cells[assumedMaxRowsCount, primaryExcelColumnIndex].Value != null)
            //                throw new Exception(string.Format("The assumed max rows {0} is not enough. Please make it bigger...", assumedMaxRowsCount));

            //            if (CurrentActiveExcelSheet.Cells[rowIndex, primaryExcelColumnIndex].Value != null)
            //            {
            //                actualExcelRowCountWithoutHeader = rowIndex - 1;
            //                break;
            //            }
            //        }
            //    }

            //    // Extracts the value from Excel.
            //    if (actualExcelRowCountWithoutHeader > 0)
            //    {
            //        for (int rowIndex = Constants.HEADER_ROW_INDEX + 1; rowIndex <= actualExcelRowCountWithoutHeader + 1; rowIndex++)
            //        {
            //            DataRow dr = excelTable.NewRow();
            //            for (int colIndex = 1; colIndex <= actualExcelColCount; colIndex++)
            //            {
            //                string columnName = CurrentActiveExcelSheet.Cells[Constants.HEADER_ROW_INDEX, colIndex].Value;
            //                dr[columnName] = CurrentActiveExcelSheet.Cells[rowIndex, colIndex].Value;
            //                //dr[colIndex - 1] = CurrentActiveExcelSheet.Cells[rowIndex, colIndex].Value;
            //            }
            //            if (OnReadingExcelRow != null)
            //                OnReadingExcelRow(this, new DataRowEventArgs(dr, (primaryExcelColumnIndex - 1).ToString()));

            //            excelTable.Rows.Add(dr);
            //        }
            //    }

            //    _excelSet.Tables.Add(excelTable);
            //}
            //XmlDocument xdoc = new XmlDocument();
            //xdoc.Load("rcg_post_temp.xml");
            //StreamReader sr = new StreamReader("rcg_post_temp.xml");
            if (!File.Exists(filename + ".xsd")) // If the previous schema did not exist, we use the meatadata to generate.
            {
                foreach (SheetConfig sheetConfig in _config.Sheets)
                {
                    if (!sheetConfig.Enabled)
                        continue;

                    DataTable metadataTable = Utility.FindMetadataTable(_metadataSet, sheetConfig.Name);
                    SetupOutputTableSchema(metadataTable, sheetConfig);
                }
                _metadataSet.WriteXmlSchema(filename + ".xsd");
                //_metadataSet.Clear();
            }
            if (File.Exists(filename + ".xsd"))
            {
                _excelSet.ReadXmlSchema(filename + ".xsd");
            }
            // TODO: Current issue is, I could not read the use the serialize and deserialise to the xml correctly.
            if (File.Exists(filename))
            {
                _excelSet.ReadXml(filename);
            }
            
        }

        private static void SetupOutputTableSchema(DataTable metadataTable, SheetConfig sheetConfig)
        {
            foreach (ColumnConfig column in sheetConfig.Columns)
            {
                if (!column.Enabled || metadataTable.Columns.Contains(column.DisplayName))
                    continue;
                metadataTable.Columns.Add(new DataColumn(column.DisplayName));
            }
        }

        public void ProcessMetadataTable()
        {
            foreach (SheetConfig sheetConfig in _config.Sheets)
            {
                if (!sheetConfig.Enabled)
                    continue;

                CurrentSheetConfig = sheetConfig;
                //CurrentActiveExcelSheet = ExcelOperationWrapper.FindExcelActiveSheet(_excel, sheetConfig.Name);

                // Fill output table content by reading row by row from metadata.
                DataTable metadataTable = Utility.FindMetadataTable(_metadataSet, sheetConfig.Name);
                SetupOutputTableSchema(metadataTable, sheetConfig);

                foreach (DataRow metadataRow in metadataTable.Rows)
                {
                    #region Process metadata
                    if (OnProcessingMetadata != null)
                        OnProcessingMetadata(this, new DataRowEventArgs(metadataRow, Constants.COLUMN_Path));

                    foreach (DataColumn dcOutput in metadataTable.Columns)
                    {
                        if (Utility.IsExtractFromMetadata(dcOutput.ColumnName))
                            continue;
                        // Get metadata content.
                        string originalContent = string.Empty;
                        ColumnConfig columnConfig = FindColumnConfig(_config, metadataTable.TableName, dcOutput.ColumnName);
                        
                        CurrentColumnConfig = columnConfig;

                        // The column with Enabled=false couldn't even be added. So we no need write the code to skip the disabled column.
                        if (!Utility.IsValidExtractFrom(columnConfig.ExtractFrom.Trim()))
                            continue;

                        originalContent = Utility.GetDataRowContent(metadataRow, columnConfig.ExtractFrom.Trim());

                        // Process the metadata.
                        BaseRuleProcessor rp = RuleProcessorFactory.GetRuleProcessor(columnConfig, this);
                        string procssedContent = rp.Process(originalContent);

                        metadataRow[dcOutput] = procssedContent;
                    }
                    #endregion

                    #region Set row mode
                    if (OnSettingRowMode != null)
                        OnSettingRowMode(this, new DataRowEventArgs(metadataRow, Constants.COLUMN_Path));

                    metadataRow[Constants.COLUMN_RowMode] = Constants.ROW_MODE_Ignored;

                    if (CurrentSheetConfig.Mode == Constants.SHEET_MODE_Refresh)
                        metadataRow[Constants.COLUMN_RowMode] = Constants.ROW_MODE_Refresh;
                    else
                    {
                        DataRowExistsOrExpires r = IsDataRowExistsOrExpires(metadataRow);

                        bool isRowExists = (r != DataRowExistsOrExpires.NotExists);

                        if (!isRowExists)
                            metadataRow[Constants.COLUMN_RowMode] = Constants.ROW_MODE_Append;
                        else
                        {
                            bool isRowExpires = (r == DataRowExistsOrExpires.ExistsAndExpires);
                            if (isRowExpires)
                                metadataRow[Constants.COLUMN_RowMode] = Constants.ROW_MODE_Update;
                        }
                    }

                    #endregion

                    #region Filter
                    if (OnFiltering != null)
                        OnFiltering(this, new DataRowEventArgs(metadataRow, Constants.COLUMN_Path));

                    bool filtered = false;
                    foreach (FilterConfig filterConfig in sheetConfig.Filters)
                    {
                        if (!filterConfig.Enabled)
                            continue;
                        IFilter filter = FilterFactory.GetFilter(filterConfig.RuleType, filterConfig.Rule, this);
                        if (filter.Match(Utility.GetDataRowContent(metadataRow, filterConfig.ExtractFrom.Trim())))
                        {
                            //if (this.OnFiltering != null)
                            //    OnFiltering(this, new OnFilteringEventArgs(GetDataRowContent(metadataRow, filterConfig.ExtractFrom.Trim()), metadataRow));
                            filtered = true;
                            break;
                        }
                    }
                    metadataRow[Constants.COLUMN_FilterFlag] = filtered;
                    if (filtered)
                        metadataRow[Constants.COLUMN_RowMode] = Constants.ROW_MODE_Filtered;

                    #endregion

                    #region Formatter
                    if (OnFormatting != null)
                        OnFormatting(this, new DataRowEventArgs(metadataRow, Constants.COLUMN_Path));

                    foreach (FormatterConfig formatterConfig in sheetConfig.Formatters)
                    {
                        if (!formatterConfig.Enabled)
                            continue;
                        //We should instanc the formatter when use it, not here.
                        IFormatter formatter = FormatterFactory.GetFormatter(formatterConfig.RuleType, formatterConfig.Rule, formatterConfig.FormatString, this);
                        string formatterToken = formatterConfig.Token;
                        // Same issue with above, we should put off to decide when to apply formatter.
                        if (formatter.Match(metadataRow, formatterConfig))
                        {
                            metadataRow[Constants.COLUMN_Formatter] = formatterToken;
                        }
                    }
                    #endregion
                }
            }
        }

        private void InitExcelActiveSheet()
        {
            // Inits the active sheet.
            // Baseline exists, then use it. ** Baseline concept is used for the original
            // design that load the excel replaced by loading the xml. So now, this concept is 
            // not that useful.
            //if (!string.IsNullOrEmpty(_config.BaselinePath.Trim()) &&
            //    File.Exists(_config.BaselinePath.Trim()))
            //{
            //    _excel.Application.Workbooks.Open(_config.BaselinePath.Trim());
            //}
            // Output exists but baseline not, use output.
            if (!string.IsNullOrEmpty(_config.OutputPath.Trim()) &&
                File.Exists(_config.OutputPath.Trim()))
            {
                if (_config.Backup)
                    File.Copy(_config.OutputPath, _config.OutputPath + DateTime.Now.ToString("yyyyMMddHHmmss"));
                bool deleted = false;
                while (!deleted)
                {
                    try
                    {
                        File.Delete(_config.OutputPath.Trim());
                        deleted = true;
                    }
                    catch (IOException ioex)
                    {
                        Console.WriteLine("Error: " + ioex.Message + " Press any key when ready...");
                        Console.Read();
                    }
                }
            }
            // Neither baseline nor output exists, use template.
            if (!string.IsNullOrEmpty(_config.TemplatePath.Trim()) &&
                File.Exists(_config.TemplatePath.Trim()))
                _excel.Application.Workbooks.Open(_config.TemplatePath.Trim());
            else
                _excel.Application.Workbooks.Add(true);
        }
        public void RefreshExcel()
        {
            _excel = new Excel.Application();
            InitExcelActiveSheet();
            try
            {
                foreach (DataTable table in _metadataSet.Tables)
                {
                    SheetConfig sheetConfig = FindSheetConfig(_config, table.TableName);

                    CurrentActiveExcelSheet = ExcelOperationWrapper.FindExcelActiveSheet(_excel, sheetConfig.Name);

                    // Generates header
                    int excelColIndex = 0;
                    foreach (DataColumn col in table.Columns)
                    {
                        if (Utility.IsExtractFromMetadata(col.ColumnName))
                            continue;
                        if (!Utility.IsColumnToOutput(col.ColumnName, CurrentSheetConfig))
                            continue;

                        excelColIndex++;
                        CurrentActiveExcelSheet.Cells[Constants.HEADER_ROW_INDEX, excelColIndex] = col.ColumnName;
                    }

                    // Clears the excel sheet while mode is "refersh"
                    if (sheetConfig.Mode == Constants.SHEET_MODE_Refresh)
                        ExcelOperationWrapper.ClearExcelSheetWithoutHeader(CurrentActiveExcelSheet, GetAvailableExcelRowCountWithoutHeader());

                    // Generates rows
                    int rowIndexToWrite = 2;
                    int existedExcelRowsCount = GetAvailableExcelRowCountWithoutHeader();
                    ExcelOperationWrapper.ClearExcelSheetFormatWithoutHeader(CurrentActiveExcelSheet, existedExcelRowsCount);
                    //Collection<DataRow> updateRowCollection = new Collection<DataRow>();
                    //Dictionary<int, DataRow> updateRowDict = new Dictionary<int, DataRow>();
                    foreach (DataRow row in table.Rows)
                    {
                        string rowMode = (string)row[Constants.COLUMN_RowMode];

                        if (rowMode == Constants.ROW_MODE_Filtered)
                            continue;

                        if (OnWritingDataRow != null)
                            OnWritingDataRow(this, new DataRowEventArgs(row, Constants.COLUMN_Path));

                        // First add all the rows which mode is "append" and marked all rows which mode is "update".
                        // Second update all the rows which mode is "update".
                        if (rowMode == Constants.ROW_MODE_Ignored)
                        {
                            if (WriteDataRow(row, CurrentActiveExcelSheet, rowIndexToWrite))
                                rowIndexToWrite++;
                        }
                        else if (rowMode == Constants.ROW_MODE_Append)
                        {
                            if (rowIndexToWrite == 0)
                                rowIndexToWrite = existedExcelRowsCount + 2;
                            if (WriteDataRow(row, CurrentActiveExcelSheet, rowIndexToWrite))
                                rowIndexToWrite++;
                        }
                        else if (rowMode == Constants.ROW_MODE_Update)
                        {
                            //updateRowCollection.Add(row);
                            // Should reserve the 'rowIndexToWrite' for update.
                            //updateRowDict[rowIndexToWrite] = row;
                            WriteDataRow(row, CurrentActiveExcelSheet, rowIndexToWrite);
                            rowIndexToWrite++;
                        }
                        else if (rowMode == Constants.ROW_MODE_Refresh)
                        {
                            if (rowIndexToWrite == 0)
                                rowIndexToWrite = 2;
                            if (WriteDataRow(row, CurrentActiveExcelSheet, rowIndexToWrite))
                                rowIndexToWrite++;
                        }
                    }
                    //foreach (var pair in updateRowDict)
                    //{
                    //    WriteDataRow(pair.Value, CurrentActiveExcelSheet, pair.Key);
                    //}
                    //foreach (DataRow row in updateRowCollection)
                    //{
                    //    int absoluteRowIndexToWrite = GetExcelRowIndex(row) + 2;
                    //    if (WriteDataRow(row, CurrentActiveExcelSheet, absoluteRowIndexToWrite))
                    //        rowIndexToWrite = GetAvailableExcelRowCountWithoutHeader() + 2;
                    //}
                }

                _excel.ActiveWorkbook.SaveAs(_config.OutputPath);
            }
            finally
            {
                _excel.Quit();
                _excel = null;

                GC.Collect();
            }
        }

        #endregion

        #region Instance Private Methods

        // 0=Unknown
        // 1=Not Exists
        // 2=Exists but Not Expires
        // 3=Exists and Expires.
        internal DataRowExistsOrExpires IsDataRowExistsOrExpires(DataRow dr)
        {
            int primaryColumnIndexOfDatatable = GetPrimaryDataColumnIndex();
            DataRowExistsOrExpires dree = Cache.Instance.GetDataRowExistsOrExpiresDictCacheValue(CurrentSheetConfig.Name, (string)dr[primaryColumnIndexOfDatatable]);
            if (dree == DataRowExistsOrExpires.UnKnown)
            {
                int rowIndexOfExcel = GetExcelRowIndex(dr);
                if (rowIndexOfExcel == Constants.INT_NOT_FOUND_INDEX)
                {
                    Cache.Instance.SetDataRowExistsOrExpiresDictCacheValue(CurrentSheetConfig.Name, (string)dr[primaryColumnIndexOfDatatable], DataRowExistsOrExpires.NotExists);
                    return DataRowExistsOrExpires.NotExists; // Not Exists
                }

                int timestampColumnIndexOfExcel = GetTimestampExcelColumnIndex();
                int timestampColumnIndexOfDatatable = GetTimestampDataColumnIndex();

                DateTime dtOfExcel = DateTime.Parse((string)Utility.FindExcelTable(_excelSet, CurrentSheetConfig.Name).Rows[rowIndexOfExcel][timestampColumnIndexOfExcel]);
                DateTime dtOfDatatable = DateTime.Parse(dr[timestampColumnIndexOfDatatable].ToString());

                bool isExpires = dtOfDatatable > dtOfExcel;
                if (isExpires)
                {
                    Cache.Instance.SetDataRowExistsOrExpiresDictCacheValue(CurrentSheetConfig.Name, (string)dr[primaryColumnIndexOfDatatable], DataRowExistsOrExpires.ExistsAndExpires);
                    return DataRowExistsOrExpires.ExistsAndExpires; // Exists and Expires
                }
                else
                {
                    Cache.Instance.SetDataRowExistsOrExpiresDictCacheValue(CurrentSheetConfig.Name, (string)dr[primaryColumnIndexOfDatatable], DataRowExistsOrExpires.ExistsButNotExpires);
                    return DataRowExistsOrExpires.ExistsButNotExpires; // Exists but Not Expires
                }
            }
            return dree;
        }

        internal int GetExcelRowIndex(DataRow dr)
        {
            int primaryColumnIndexOfDatatable = GetPrimaryDataColumnIndex();
            int primaryColumnIndexOfExcel = GetPrimaryExcelColumnIndex();

            int rowIndex = 0;
            foreach (DataRow row in Utility.FindExcelTable(_excelSet, CurrentSheetConfig.Name).Rows)
            {
                if ((string)row[primaryColumnIndexOfExcel] == (string)dr[primaryColumnIndexOfDatatable])
                    return rowIndex;
                rowIndex++;
            }
            return Constants.INT_NOT_FOUND_INDEX;
        }

        internal int GetTimestampDataColumnIndex()
        {
            int timestampDataColumnIndex = Cache.Instance.GetCacheValue<int>(CurrentSheetConfig.Name, Cache.Timestamp_DataColumn_Index);
            if (timestampDataColumnIndex <= 0)
            {
                foreach (ColumnConfig columnConfig in CurrentSheetConfig.Columns)
                {
                    if (!columnConfig.Timestamp)
                        continue;
                    DataTable table = Utility.FindMetadataTable(_metadataSet, CurrentSheetConfig.Name);
                    int index = 0;
                    foreach (DataColumn dc in table.Columns)
                    {
                        if (dc.ColumnName == columnConfig.DisplayName)
                        {
                            Cache.Instance.SetCacheValue<int>(CurrentSheetConfig.Name, Cache.Timestamp_DataColumn_Index, index);
                            return index;
                        }
                        index++;
                    }
                }
            }
            return timestampDataColumnIndex;
        }

        internal int GetPrimaryDataColumnIndex()
        {
            int primaryDataColumnIndex = Cache.Instance.GetCacheValue<int>(CurrentSheetConfig.Name, Cache.Primary_DataColumn_Index);
            if (primaryDataColumnIndex <= 0)
            {
                foreach (ColumnConfig columnConfig in CurrentSheetConfig.Columns)
                {
                    if (!columnConfig.Primary)
                        continue;
                    DataTable table = Utility.FindMetadataTable(_metadataSet, CurrentSheetConfig.Name);
                    int index = 0;
                    foreach (DataColumn dc in table.Columns)
                    {
                        if (dc.ColumnName == columnConfig.DisplayName)
                        {
                            Cache.Instance.SetCacheValue<int>(CurrentSheetConfig.Name, Cache.Primary_DataColumn_Index, index);
                            return index;
                        }
                        index++;
                    }
                }
            }
            return primaryDataColumnIndex;
        }

        internal int GetTimestampExcelColumnIndex()
        {
            int timestampExcelColumnIndex = Cache.Instance.GetCacheValue<int>(CurrentSheetConfig.Name, Cache.Timestamp_ExcelColumn_Index);
            if (timestampExcelColumnIndex <= 0)
            {
                foreach (ColumnConfig columnConfig in CurrentSheetConfig.Columns)
                {
                    if (!columnConfig.Timestamp)
                        continue;

                    int excelColIndex = 0;
                    foreach (DataColumn dc in Utility.FindExcelTable(_excelSet, CurrentSheetConfig.Name).Columns)
                    {
                        if (dc.ColumnName == columnConfig.DisplayName)
                        {
                            Cache.Instance.SetCacheValue<int>(CurrentSheetConfig.Name, Cache.Timestamp_ExcelColumn_Index, excelColIndex);
                            return excelColIndex;
                        }
                        excelColIndex++;
                    }
                }
            }
            return timestampExcelColumnIndex;
        }

        internal int GetPrimaryExcelColumnIndex()
        {
            int primaryExcelColumnIndex = Cache.Instance.GetCacheValue<int>(CurrentSheetConfig.Name, Cache.Primary_ExcelColumn_Index);
            if (primaryExcelColumnIndex <= 0)
            {
                foreach (ColumnConfig columnConfig in CurrentSheetConfig.Columns)
                {
                    if (!columnConfig.Primary)
                        continue;

                    int excelColIndex = 0;
                    foreach (DataColumn dc in Utility.FindExcelTable(_excelSet, CurrentSheetConfig.Name).Columns)
                    {
                        if (dc.ColumnName == columnConfig.DisplayName)
                        {
                            Cache.Instance.SetCacheValue<int>(CurrentSheetConfig.Name, Cache.Primary_ExcelColumn_Index, excelColIndex);
                            return excelColIndex;
                        }
                        excelColIndex++;
                    }
                }
            }
            return primaryExcelColumnIndex;
        }

        internal int GetAvailableExcelRowCountWithoutHeader()
        {
            return Utility.FindExcelTable(_excelSet, CurrentSheetConfig.Name).Rows.Count;
        }

        internal int GetAvailableExcelColumnCount()
        {
            return Utility.FindExcelTable(_excelSet, CurrentSheetConfig.Name).Columns.Count;
        }

        private bool WriteDataRow(DataRow row, dynamic activeSheet, int rowIndexToWrite)
        {
            int realExcelRowIndex = rowIndexToWrite;

            #region Apply Formatter
            string formatterToken = row[Constants.COLUMN_Formatter] as string;
            if (!string.IsNullOrEmpty(formatterToken))
            {
                //IFormatter formatter = row[Constants.COLUMN_Formatter] as IFormatter;
                string[] formatterTokens = formatterToken.Split(new string[] { FormatterConfig.TokenSplitter }, StringSplitOptions.None);
                IFormatter formatter = FormatterFactory.GetFormatter(formatterTokens[0], formatterTokens[1], formatterTokens[2], this);
                if (formatter != null)
                    formatter.Execute(realExcelRowIndex);
            }
            #endregion

            int excelColIndex = 0;
            foreach (DataColumn col in row.Table.Columns)
            {
                if (Utility.IsExtractFromMetadata(col.ColumnName))
                    continue;
                if (!Utility.IsColumnToOutput(col.ColumnName, CurrentSheetConfig))
                    continue;

                excelColIndex++;
                ColumnConfig columnConfig = FindColumnConfig(_config, row.Table.TableName, col.ColumnName);
                if (columnConfig.ExtractFrom == Constants.PREDEFINED_AutoIncrease)
                    activeSheet.Cells[realExcelRowIndex, excelColIndex] = (realExcelRowIndex - 1).ToString();
                else
                    activeSheet.Cells[realExcelRowIndex, excelColIndex] = row[col.ColumnName].ToString();
            }

            return true;
        }

        #endregion

        #region Events

        public event EventHandler<HandlableExceptionEventArgs> OnHandlableException;

        public event EventHandler<DataRowEventArgs> OnReadingMetadata;

        public event EventHandler<DataRowEventArgs> OnReadingExcelRow;

        public event EventHandler<DataRowEventArgs> OnProcessingMetadata;

        public event EventHandler<DataRowEventArgs> OnSettingRowMode;

        public event EventHandler<DataRowEventArgs> OnFiltering;

        public event EventHandler<DataRowEventArgs> OnFormatting;

        public event EventHandler<DataRowEventArgs> OnWritingDataRow;

        #endregion
    }

    public class HandlableExceptionEventArgs : EventArgs
    {
        public Exception HandlableException { get; private set; }
        public string KeyMessage { get; private set; }

        public HandlableExceptionEventArgs(Exception handlableException, string keyMessage)
        {
            this.HandlableException = handlableException;
            this.KeyMessage = keyMessage;
        }
    }

    public class DataRowEventArgs : EventArgs
    {
        public DataRow RowProcessing { get; private set; }
        public string KeyMessage { get; private set; }

        public DataRowEventArgs(DataRow row, string keyMessage)
        {
            this.RowProcessing = row;
            this.KeyMessage = keyMessage;
        }
    }
}
