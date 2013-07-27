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
        private DataSet _combinedSet = null;
        private Excel.Application _excel;
        public List<string> _metadataPrimaryColumnList = new List<string>();
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
            return null;
            //throw new Exception(string.Format("Sheet config {0} not recognized...", sheetName));
        }

        private ColumnConfig FindColumnConfig(ConfigDoc config, string sheetName, string columnName)
        {
            SheetConfig sheetConfig = FindSheetConfig(config, sheetName);
            if (sheetConfig == null)
                return null;

            foreach (ColumnConfig columnConfig in sheetConfig.Columns)
            {
                if (columnConfig.Name == columnName)
                {
                    CurrentColumnConfig = columnConfig;
                    return columnConfig;
                }
            }
            return null;
            //throw new Exception(string.Format("Column config {0} of Sheet {1} is not recognized...", columnName, sheetName));
        }

        #endregion

        #region Instance Methods

        /// <summary>
        /// Reads the configuration within "Mappings.xml"
        /// </summary>
        /// <param name="configFileName"></param>
        public void ReadConfiguration(string configFileName, string hijackConfigurationParameter, bool onlyAttributeIfHijack)
        {
            _config.Read(configFileName, hijackConfigurationParameter, onlyAttributeIfHijack);
            //if (!string.IsNullOrEmpty(hijackConfigurationParameter))
            //{
            //    this.HijackConfiguration(hijackConfigurationParameter);
            //}
            _config.Validate();
        }

        /// <summary>
        /// Reads the data which resideds in your harddisk and generated metadata based on it.
        /// </summary>
        public void GenerateMetadata()
        {
            _metadataPrimaryColumnList.Clear();
            foreach (SheetConfig sheetConfig in _config.Sheets)
            {
                if (!sheetConfig.Enabled)
                    continue;

                CurrentSheetConfig = sheetConfig;

                DataTable dt = new DataTable(sheetConfig.Name);
                // Use the "Prefix" to save the Mode.
                dt.Prefix = sheetConfig.Mode;

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
                DataColumn dcPrimaryColumnIndex = new DataColumn(Constants.COLUMN_PrimaryColumnIndex, typeof(int));
                dt.Columns.Add(dcPrimaryColumnIndex);
                DataColumn dcTimestampColumnIndex = new DataColumn(Constants.COLUMN_TimestampColumnIndex, typeof(int));
                dt.Columns.Add(dcTimestampColumnIndex);
                DataColumn dcOutputColumnIndex = new DataColumn(Constants.COLUMN_OutputColumnIndex, typeof(string));
                dt.Columns.Add(dcOutputColumnIndex);
                DataColumn dcHyperlinkColumnIndex = new DataColumn(Constants.COLUMN_HyperlinkColumnIndex, typeof(string));
                dt.Columns.Add(dcHyperlinkColumnIndex);
                DataColumn dcAutoIncreaseColumnIndex = new DataColumn(Constants.COLUMN_AutoIncreaseColumnIndex, typeof(string));
                dt.Columns.Add(dcAutoIncreaseColumnIndex);
                DataColumn dcLocationFrom = new DataColumn(Constants.COLUMN_LocationFrom, typeof(string));
                dt.Columns.Add(dcLocationFrom);

                SetupOutputTableSchema(dt, sheetConfig);
                // Read metadata.
                foreach (LocationConfig locationConfig in sheetConfig.Locations)
                {
                    if (!locationConfig.Enabled) continue;
                    if (!Utility.IsLocationAccessable(locationConfig.Path))
                    {
                        if (OnGeneralMessageException != null)
                            OnGeneralMessageException(this, new GeneralMessageEventArgs(string.Format("Location {0} is not accessible...", locationConfig.Path)));
                        continue;
                    }
                    //try
                    //{
                    ReadMetadata(locationConfig, dt);
                    //}
                    //catch (IOException ex) // This is redundant since the "IsLocationAccessable" method added. 
                    //{
                    //    if (OnHandlableException != null)
                    //        OnHandlableException(this, new HandlableExceptionEventArgs(ex, string.Format("Error when reading metadata", locationConfig.Path)));
                    //}
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

            // Checks whether the path is valid or not. If not exists, just skip it, since 
            if (string.IsNullOrEmpty(sPath) || !Directory.Exists(sPath))
                return;

            // Include folder indicates the metadata generation is based on Folder.
            if (locationConfig.IncludeFolder)
            {
                dirList = Directory.GetDirectories(sPath);
            }
            // The metadata generationis based on file.
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
            // Read metadata
            foreach (string dir in dirList)
            {
                try
                {
                    DirectoryInfo d = new DirectoryInfo(dir);
                    DataRow dr = dt.NewRow();
                    dr[Constants.COLUMN_Path] = dir;
                    dr[Constants.COLUMN_LastModified] = d.LastWriteTime.ToString();
                    dr[Constants.COLUMN_Attributes] = d.Attributes.ToString();
                    dr[Constants.COLUMN_LocationFrom] = locationConfig.Name;

                    long size = 0;
                    int fileCount = 0;
                    int subFolderCount = 0;
                    string filesType = string.Empty;
                    try
                    {
                        Utility.CalFolderSize(ref size, ref fileCount, ref subFolderCount, ref filesType, d);
                    }
                    catch (UnauthorizedAccessException)
                    {
                        size = -1;
                        fileCount = -1;
                        subFolderCount = -1;
                        filesType = "Access Denied";
                    }
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

                    // Append to list here is added for new check deleted item function. ==>
                    int columnIndex = -1;
                    foreach (DataColumn dc in dt.Columns)
                    {
                        columnIndex++;
                        if (Utility.IsExtractFromMetadata(dc.ColumnName)) // The column with pattern "_..._" is invisible to user, so all the special columns should be intened to actual columns. So we skip the pattern "_..._" here.
                            continue;
                        ColumnConfig columnConfig = FindColumnConfig(_config, dt.TableName, dc.ColumnName);
                        SetSpecialColumnIndex(columnConfig, columnIndex, dr);
                    }

                    dt.Rows.Add(dr);

                    AppendToList(dr, _metadataPrimaryColumnList);
                    // <==
                }
                catch (IOException ex)
                {
                    if (OnHandlableException != null)
                        OnHandlableException(this, new HandlableExceptionEventArgs(ex, 
                            string.Format("Error when reading metadata... {0}", dir) + 
                            Environment.NewLine + 
                            ex.ToString() +
                            "Press any key to continue..."));
                    Console.Read();
                }
            }
        }

        private void AppendToList(DataRow dr, List<string> list)
        {
            list.Add((string)dr[Constants.COLUMN_Path]);
        }

        private DataSet CombineMetadataExcelDataset()
        {
            if (_combinedSet == null)
            {
                DataSet datasetToExecute = new DataSet();

                foreach (DataTable table in _metadataSet.Tables)
                {
                    if (table.Prefix == Constants.SHEET_MODE_Refresh)
                    {
                        // If the refresh mode, we could abandon the exising data in excel dataset.
                        if (!datasetToExecute.Tables.Contains(table.TableName))
                            datasetToExecute.Tables.Add(table.Copy());
                    }
                    else
                    {
                        // Merge the metadata into existing table.
                        DataTable tableMetadata = Utility.FindMetadataTable(_metadataSet, table.TableName);
                        DataTable tableExcel = Utility.FindExcelTable(_excelSet, table.TableName);
                        if (tableExcel != null)
                        {
                            // Counts the non-filter rows in existing table ==>
                            int nonFilteredRowsCount = 0;
                            foreach (DataRow dr in tableExcel.Rows)
                            {
                                #region Formatter column should be erased excepet deleted formatter
                                string ruleType, rule, formatString;
                                FormatterConfig.SplitRuleType_Rule_Formatter((string)dr[Constants.COLUMN_Formatter], out ruleType, out rule, out formatString);
                                if (rule != GetDeletedItemFormatter(table.TableName).Rule)
                                    dr[Constants.COLUMN_Formatter] = string.Empty;
                                #endregion

                                if (!(bool)dr[Constants.COLUMN_FilterFlag])
                                    nonFilteredRowsCount++;
                            }
                            // <==

                            if (tableMetadata != null)
                            {
                                string[] autoIncreasedColumnIndexes = ((string)FindFirstUnfilteredRow(tableMetadata)[Constants.COLUMN_AutoIncreaseColumnIndex]).Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                int rowIndex = nonFilteredRowsCount;
                                foreach (DataRow dr in tableMetadata.Rows)
                                {
                                    if ((string)dr[Constants.COLUMN_RowMode] == Constants.ROW_MODE_Ignored)
                                        continue;

                                    // Should not increase AutoNumber when mode is "filetered".
                                    if ((string)dr[Constants.COLUMN_RowMode] == Constants.ROW_MODE_Filtered)
                                        continue;

                                    if ((string)dr[Constants.COLUMN_RowMode] == Constants.ROW_MODE_Update)
                                    {
                                        DataRow sameRow = Utility.FindSameRow(dr, tableExcel);
                                        Dictionary<int, int> keptAutoIncreasedIndexes = GetAllAutoIncreasedIndexesValue(sameRow, autoIncreasedColumnIndexes);
                                        int index = tableExcel.Rows.IndexOf(sameRow);
                                        tableExcel.Rows.RemoveAt(index);
                                        DataRow newRow2 = tableExcel.NewRow();
                                        #region Set auto increased column & value
                                        for (int i = 0; i < tableMetadata.Columns.Count; i++)
                                        {
                                            newRow2[i] = dr[i];// copy the value to new row.
                                        }
                                        foreach (KeyValuePair<int, int> kvp in keptAutoIncreasedIndexes)
                                        {
                                            newRow2[kvp.Key] = kvp.Value;
                                        }
                                        #endregion
                                        tableExcel.Rows.InsertAt(newRow2, index);
                                    }
                                    else
                                    {
                                        rowIndex++;
                                        DataRow newRow = tableExcel.NewRow();

                                        #region Special columns that should not be moved from existing table directly

                                        #region Set auto increased column & value
                                        for (int i = 0; i < tableMetadata.Columns.Count; i++)
                                        {

                                            foreach (string index in autoIncreasedColumnIndexes)
                                            {
                                                if (int.Parse(index) == i) // if is the auto increased column.
                                                    newRow[i] = rowIndex;
                                                else
                                                    newRow[i] = dr[i]; // else just copy the original value.
                                            }
                                        }
                                        #endregion

                                        #endregion

                                        tableExcel.Rows.Add(newRow);
                                    }
                                }
                            }
                            datasetToExecute.Tables.Add(tableExcel.Copy());
                        }
                        else
                        {
                            datasetToExecute.Tables.Add(tableMetadata.Copy());
                        }
                    }
                }
                foreach (DataTable table in _excelSet.Tables)
                {
                    if (!datasetToExecute.Tables.Contains(table.TableName))
                        datasetToExecute.Tables.Add(table.Copy());
                }
                _combinedSet = datasetToExecute;
            }
            return _combinedSet;
        }

        private Dictionary<int, int> GetAllAutoIncreasedIndexesValue(DataRow sameRow, string[] autoIncreasedColumnIndexes)
        {
            Dictionary<int, int> rt = new Dictionary<int, int>(); // Dictionary<ColumnIndex, Value>
            for (int i = 0; i < sameRow.Table.Columns.Count; i++)
            {
                foreach (string index in autoIncreasedColumnIndexes)
                {
                    if (int.Parse(index) == i)
                        rt[int.Parse(index)] = int.Parse((string)sameRow[i]);
                        //rt.Add(int.Parse((string)sameRow[i]));
                }
            }
            return rt;
        }

        public void OutputTemporaryFiles(string filename)
        {
            DataSet datasetToExecute = CombineMetadataExcelDataset();

            foreach (DataTable dt in datasetToExecute.Tables)
            {
                dt.WriteXmlSchema(filename + "." + dt.TableName + ".xsd");
                dt.WriteXml(filename + "." + dt.TableName + ".xml");
            }
            //datasetToExecute.WriteXml(filename);
        }
        /// <summary>
        /// Reads the original existing excel to determine what already exists.
        /// </summary>
        public void ReadPreviousMetadata(string filename)
        {
            // Output the xsd for each enabled sheet by metadata since the original xsd file cannot find.
            foreach (SheetConfig sheetConfig in _config.Sheets)
            {
                if (!sheetConfig.Enabled)
                    continue;

                //if (!File.Exists(sheetConfig.Name + ".xsd"))
                if (!File.Exists(filename + "." + sheetConfig.Name + ".xsd"))
                {
                    DataTable metadataTable = Utility.FindMetadataTable(_metadataSet, sheetConfig.Name);

                    metadataTable.WriteXmlSchema(filename + "." + sheetConfig.Name + ".xsd");
                }
            }

            // Then read the xsd files in, also by sheet basis.
            foreach (SheetConfig sheetConfig in _config.Sheets)
            {
                // As long as there is xsd file exists, we shall read them in no matter it is enabled or not.
                if (!sheetConfig.RefMode && !sheetConfig.Enabled)
                    continue;

                string schemaFilePath = filename + "." + sheetConfig.Name + ".xsd";
                if (File.Exists(schemaFilePath))
                {
                    // Read schema.
                    DataTable excelDt = new DataTable();
                    excelDt.ReadXmlSchema(schemaFilePath);
                    // Read content.
                    string contentFilePath = filename + "." + sheetConfig.Name + ".xml";
                    if (File.Exists(contentFilePath))
                    {
                        excelDt.ReadXml(contentFilePath);
                    }

                    _excelSet.Tables.Add(excelDt);
                }
            }
           
            // TODO: Current issue is, I could not read the use the serialize and deserialise to the xml correctly.
            // Reads the content in.
            //if (File.Exists(filename))
            //{
            //    DataTable excelDt = new DataTable();
            //    excelDt.ReadXml(
            //    _excelSet.ReadXml(filename);
            //}
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

                int rowIndex = -1;
                int autoIncreaseNumber = 0;

                // Check deleted status ==>
                foreach (LocationConfig locationConfig in sheetConfig.Locations)
                {
                    if (!locationConfig.Enabled) continue;
                    if (!Utility.IsLocationAccessable(locationConfig.Path)) continue;
                    foreach (DataRow dr in _excelSet.Tables[sheetConfig.Name].Rows)
                    {
                        if (string.Compare((string)dr[Constants.COLUMN_LocationFrom], locationConfig.Name.Trim(), true) == 0)
                        {
                            if (!_metadataPrimaryColumnList.Contains((string)dr[Constants.COLUMN_Path]))
                            {
                                dr[Constants.COLUMN_RowMode] = Constants.ROW_MODE_Deleted;
                                // Please refere to the comment to "GetDeletedItemFormatter"
                                dr[Constants.COLUMN_Formatter] = GetDeletedItemFormatter(sheetConfig.Name).Token;
                            }
                        }
                    }
                }
                // <==
                foreach (DataRow metadataRow in metadataTable.Rows)
                {
                    rowIndex++;

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
                    if (!filtered)
                    {
                        autoIncreaseNumber++; // auto_increase is only available for the row which is not filter out.
                        #region Process metadata
                        if (OnProcessingMetadata != null)
                            OnProcessingMetadata(this, new DataRowEventArgs(metadataRow, Constants.COLUMN_Path));
                        int columnIndex = -1;
                        foreach (DataColumn dcOutput in metadataTable.Columns)
                        {
                            columnIndex++;
                            if (Utility.IsExtractFromMetadata(dcOutput.ColumnName)) // This one is easy to explain, processing shall ignored the columns with pattern like "_..._" since they are only generated by metadata, need not process.
                                continue;
                            // Get metadata content.
                            string originalContent = string.Empty;
                            ColumnConfig columnConfig = FindColumnConfig(_config, metadataTable.TableName, dcOutput.ColumnName);

                            CurrentColumnConfig = columnConfig;

                            #region Set the special columns index
                            //SetSpecialColumnIndex(columnConfig, columnIndex, metadataRow);
                            
                            #endregion

                            // Predefined column like auto increased column shall be handled specifically here.
                            if (Utility.IsPredefinedColumn(columnConfig.ExtractFrom.Trim()))
                            {
                                metadataRow[dcOutput] = autoIncreaseNumber; // "No." column shall be 1 based, not 0 based.
                            }

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
                            {
                                metadataRow[Constants.COLUMN_RowMode] = Constants.ROW_MODE_Append;
                                metadataRow[Constants.COLUMN_Formatter] = FormatterConfig.GetDefaultAppendedItemFormatterConfig().Token; // We set default appended format here, and override it with mappings definition later.
                            }
                            else
                            {
                                bool isRowExpires = (r == DataRowExistsOrExpires.ExistsAndExpires);
                                if (isRowExpires)
                                {
                                    metadataRow[Constants.COLUMN_RowMode] = Constants.ROW_MODE_Update;
                                    metadataRow[Constants.COLUMN_Formatter] = FormatterConfig.GetDefaultUpdatedItemFormatterConfig().Token; // We set default updated format here, and override it with mappings definition later.
                                }
                            }
                        }
                        
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
        }

        // The deletedItem formatter is different from appendedItem / updatedItem formatter,
        // Since deletedItem is based on Existing(Excel) data source while appendedItem / updatedItem are based on New(Metadata) source.
        // This is the reason we wrap "GetDeletedItemFormatter" here since it cannot get user configured formatter like appendedItem / updatedItem.
        private FormatterConfig GetDeletedItemFormatter(string sheetName)
        {
            SheetConfig sheetConfig = FindSheetConfig(_config, sheetName);

            foreach (FormatterConfig formatterConfig in sheetConfig.Formatters)
            {
                if (!formatterConfig.Enabled)
                    continue;
                
                if (string.Compare(formatterConfig.Rule.Trim(), Constants.FORMATTER_Internal_DeletedItem, true) == 0)
                {
                    return formatterConfig;
                }
            }
            // If not defined in Mappings.xml, then use the default config.
            return FormatterConfig.GetDefaultDeletedItemFormatterConfig();
        }

        private static void SetSpecialColumnIndex(ColumnConfig columnConfig, int columnIndex, DataRow metadataRow)
        {
            if (columnConfig.Primary)
                metadataRow[Constants.COLUMN_PrimaryColumnIndex] = columnIndex;
            if (columnConfig.Timestamp)
                metadataRow[Constants.COLUMN_TimestampColumnIndex] = columnIndex;
            if (columnConfig.Output)
                metadataRow[Constants.COLUMN_OutputColumnIndex] += string.Format("{0},", columnIndex);
            if (columnConfig.Hyperlink)
                metadataRow[Constants.COLUMN_HyperlinkColumnIndex] += string.Format("{0},", columnIndex);
            if (columnConfig.ExtractFrom == Constants.PREDEFINED_AutoIncrease)
                metadataRow[Constants.COLUMN_AutoIncreaseColumnIndex] += string.Format("{0},", columnIndex);
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
                    catch (IOException ex)
                    {
                        if (OnHandlableException != null)
                            OnHandlableException(this, new HandlableExceptionEventArgs(ex, "Error: " + ex.Message + " Press any key when ready..."));
                        Console.Read();
                    }
                }
            }
            // Neither baseline nor output exists, use template.
            // *Baseline is obseleted.
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
            // Combine the _metadataSet & _excelSet for the case first deinfed in mappings.xml then be disabled ==>
            DataSet datasetToExecute = CombineMetadataExcelDataset();
            //DataSet datasetToExecute = _metadataSet;
            // <==
            try
            {
                foreach (DataTable table in datasetToExecute.Tables)
                {
                    CurrentActiveExcelSheet = ExcelOperationWrapper.FindExcelActiveSheet(_excel, table.TableName);

                    // Generates header
                    int excelColIndex = 0;
                    int loopColIndex = -1;
                    foreach (DataColumn col in table.Columns)
                    {
                        loopColIndex++;
                        // Some clarification here, actually the logic "IsExtractFromMetadata" does same check with "!IsColumnToOutput".
                        // So here we have some duplicated logic check, but since "IsColumnToOutput" is newly added, we just keep there to avoid unexpected bugs.
                        
                        // Yeh, after more investigation, I found there are slight differences between "IsExtractFromMetadata" with "!IsColumnToOutput"
                        // The column could defined attribute "output=false" that is not OutputColumn but not ExtractFromMetadata either
                        // This logic is correct, do NOT change.
                        if (Utility.IsExtractFromMetadata(col.ColumnName))
                            continue;

                        if (!Utility.IsColumnToOutput(FindFirstUnfilteredRow(table), loopColIndex))
                            continue;

                        excelColIndex++;
                        CurrentActiveExcelSheet.Cells[Constants.HEADER_ROW_INDEX, excelColIndex] = col.ColumnName;
                    }
                    
                    // Clears the excel sheet while mode is "refersh" - Prefix saved "Mode" value.
                    if (table.Prefix == Constants.SHEET_MODE_Refresh && _excelSet.Tables.Contains(table.TableName))
                    {
                        ExcelOperationWrapper.ClearExcelSheetWithoutHeader(CurrentActiveExcelSheet, GetAvailableExcelRowCountWithoutHeader(table.TableName));
                    }
                    // Generates rows
                    int rowIndexToWrite = 2;
                    int existedExcelRowsCount = 0;
                    if (_excelSet.Tables.Contains(table.TableName))
                    {
                        existedExcelRowsCount = GetAvailableExcelRowCountWithoutHeader(table.TableName);
                        ExcelOperationWrapper.ClearExcelSheetFormatWithoutHeader(CurrentActiveExcelSheet, existedExcelRowsCount);
                    }
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
                        else if (rowMode == Constants.ROW_MODE_Deleted)
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

        private DataRow FindFirstUnfilteredRow(DataTable table)
        {
            foreach (DataRow dr in table.Rows)
            {
                if (!(bool)dr[Constants.COLUMN_FilterFlag])
                    return dr;
            }
            return null;
        }

        #endregion

        #region Instance Private Methods

        // 0=Unknown
        // 1=Not Exists
        // 2=Exists but Not Expires
        // 3=Exists and Expires.
        internal DataRowExistsOrExpires IsDataRowExistsOrExpires(DataRow dr)
        {
            if (!_excelSet.Tables.Contains(dr.Table.TableName))
                return DataRowExistsOrExpires.NotExists;

            int primaryColumnIndexOfDatatable = GetPrimaryDataColumnIndex(dr);
            DataRowExistsOrExpires dree = Cache.Instance.GetDataRowExistsOrExpiresDictCacheValue(dr.Table.TableName, (string)dr[primaryColumnIndexOfDatatable]);
            if (dree == DataRowExistsOrExpires.UnKnown)
            {
                int rowIndexOfExcel = GetExcelRowIndex(dr);
                if (rowIndexOfExcel == Constants.INT_NOT_FOUND_INDEX)
                {
                    Cache.Instance.SetDataRowExistsOrExpiresDictCacheValue(dr.Table.TableName, (string)dr[primaryColumnIndexOfDatatable], DataRowExistsOrExpires.NotExists);
                    return DataRowExistsOrExpires.NotExists; // Not Exists
                }

                int timestampColumnIndexOfExcel = GetTimestampExcelColumnIndex(dr);
                int timestampColumnIndexOfDatatable = GetTimestampDataColumnIndex(dr);

                bool isDeletedBefore = (string)Utility.FindExcelTable(_excelSet, dr.Table.TableName).Rows[rowIndexOfExcel][Constants.COLUMN_RowMode] == Constants.ROW_MODE_Deleted;
                DateTime dtOfExcel = DateTime.Parse((string)Utility.FindExcelTable(_excelSet, dr.Table.TableName).Rows[rowIndexOfExcel][timestampColumnIndexOfExcel]);
                DateTime dtOfDatatable = DateTime.Parse(dr[timestampColumnIndexOfDatatable].ToString());

                bool isExpires = dtOfDatatable > dtOfExcel;
                if (isExpires || isDeletedBefore)
                {
                    Cache.Instance.SetDataRowExistsOrExpiresDictCacheValue(dr.Table.TableName, (string)dr[primaryColumnIndexOfDatatable], DataRowExistsOrExpires.ExistsAndExpires);
                    return DataRowExistsOrExpires.ExistsAndExpires; // Exists and Expires
                }
                else
                {
                    Cache.Instance.SetDataRowExistsOrExpiresDictCacheValue(dr.Table.TableName, (string)dr[primaryColumnIndexOfDatatable], DataRowExistsOrExpires.ExistsButNotExpires);
                    return DataRowExistsOrExpires.ExistsButNotExpires; // Exists but Not Expires
                }
            }
            return dree;
        }

        internal int GetExcelRowIndex(DataRow dr)
        {
            int primaryColumnIndexOfDatatable = GetPrimaryDataColumnIndex(dr);
            int primaryColumnIndexOfExcel = GetPrimaryExcelColumnIndex(dr);

            int rowIndex = -1;
            foreach (DataRow row in Utility.FindExcelTable(_excelSet, CurrentSheetConfig.Name).Rows)
            {
                rowIndex++;
                if ((string)row[Constants.COLUMN_RowMode] == Constants.ROW_MODE_Filtered)
                    continue;

                if ((string)row[primaryColumnIndexOfExcel] == (string)dr[primaryColumnIndexOfDatatable])
                    return rowIndex;
            }
            return Constants.INT_NOT_FOUND_INDEX;
        }

        internal int GetTimestampDataColumnIndex(DataRow dr)
        {
            return (int)dr[Constants.COLUMN_TimestampColumnIndex];
            //int timestampDataColumnIndex = Cache.Instance.GetCacheValue<int>(dr.Table.tableName, Cache.Timestamp_DataColumn_Index);
            //if (timestampDataColumnIndex <= 0)
            //{
            //    foreach (ColumnConfig columnConfig in CurrentSheetConfig.Columns)
            //    {
            //        if (!columnConfig.Timestamp)
            //            continue;
            //        DataTable table = Utility.FindMetadataTable(_metadataSet, CurrentSheetConfig.Name);
            //        int index = 0;
            //        foreach (DataColumn dc in table.Columns)
            //        {
            //            if (dc.ColumnName == columnConfig.DisplayName)
            //            {
            //                Cache.Instance.SetCacheValue<int>(CurrentSheetConfig.Name, Cache.Timestamp_DataColumn_Index, index);
            //                return index;
            //            }
            //            index++;
            //        }
            //    }
            //}
            //return timestampDataColumnIndex;
        }

        internal int GetPrimaryDataColumnIndex(DataRow dr)
        {
            return (int)dr[Constants.COLUMN_PrimaryColumnIndex];
            //int primaryDataColumnIndex = Cache.Instance.GetCacheValue<int>(CurrentSheetConfig.Name, Cache.Primary_DataColumn_Index);
            //if (primaryDataColumnIndex <= 0)
            //{
            //    foreach (ColumnConfig columnConfig in CurrentSheetConfig.Columns)
            //    {
            //        if (!columnConfig.Primary)
            //            continue;
            //        DataTable table = Utility.FindMetadataTable(_metadataSet, CurrentSheetConfig.Name);
            //        int index = 0;
            //        foreach (DataColumn dc in table.Columns)
            //        {
            //            if (dc.ColumnName == columnConfig.DisplayName)
            //            {
            //                Cache.Instance.SetCacheValue<int>(CurrentSheetConfig.Name, Cache.Primary_DataColumn_Index, index);
            //                return index;
            //            }
            //            index++;
            //        }
            //    }
            //}
            //return primaryDataColumnIndex;
        }

        internal int GetTimestampExcelColumnIndex(DataRow dr)
        {
            int timestampExcelColumnIndex = Cache.Instance.GetCacheValue<int>(dr.Table.TableName, Cache.Timestamp_ExcelColumn_Index);
            if (timestampExcelColumnIndex <= 0)
            {
                string timestampColumnName = dr.Table.Columns[(int)dr[Constants.COLUMN_TimestampColumnIndex]].ColumnName;
                int excelColIndex = 0;
                foreach (DataColumn dc in Utility.FindExcelTable(_excelSet, CurrentSheetConfig.Name).Columns)
                {
                    if (dc.ColumnName == timestampColumnName)
                    {
                        Cache.Instance.SetCacheValue<int>(CurrentSheetConfig.Name, Cache.Timestamp_ExcelColumn_Index, excelColIndex);
                        return excelColIndex;
                    }
                    excelColIndex++;
                }
                //foreach (ColumnConfig columnConfig in CurrentSheetConfig.Columns)
                //{
                //    if (!columnConfig.Timestamp)
                //        continue;

                //    int excelColIndex = 0;
                //    foreach (DataColumn dc in Utility.FindExcelTable(_excelSet, CurrentSheetConfig.Name).Columns)
                //    {
                //        if (dc.ColumnName == columnConfig.DisplayName)
                //        {
                //            Cache.Instance.SetCacheValue<int>(CurrentSheetConfig.Name, Cache.Timestamp_ExcelColumn_Index, excelColIndex);
                //            return excelColIndex;
                //        }
                //        excelColIndex++;
                //    }
                //}
            }
            return timestampExcelColumnIndex;
        }

        internal int GetPrimaryExcelColumnIndex(DataRow dr)
        {
            int primaryExcelColumnIndex = Cache.Instance.GetCacheValue<int>(CurrentSheetConfig.Name, Cache.Primary_ExcelColumn_Index);
            if (primaryExcelColumnIndex <= 0)
            {
                string primaryColumnName = dr.Table.Columns[(int)dr[Constants.COLUMN_PrimaryColumnIndex]].ColumnName;
                int excelColIndex = 0;
                foreach (DataColumn dc in Utility.FindExcelTable(_excelSet, CurrentSheetConfig.Name).Columns)
                {
                    if (dc.ColumnName == primaryColumnName)
                    {
                        Cache.Instance.SetCacheValue<int>(CurrentSheetConfig.Name, Cache.Primary_ExcelColumn_Index, excelColIndex);
                        return excelColIndex;
                    }
                    excelColIndex++;
                }
                //foreach (ColumnConfig columnConfig in CurrentSheetConfig.Columns)
                //{
                //    if (!columnConfig.Primary)
                //        continue;

                //    int excelColIndex = 0;
                //    foreach (DataColumn dc in Utility.FindExcelTable(_excelSet, CurrentSheetConfig.Name).Columns)
                //    {
                //        if (dc.ColumnName == columnConfig.DisplayName)
                //        {
                //            Cache.Instance.SetCacheValue<int>(CurrentSheetConfig.Name, Cache.Primary_ExcelColumn_Index, excelColIndex);
                //            return excelColIndex;
                //        }
                //        excelColIndex++;
                //    }
                //}
            }
            return primaryExcelColumnIndex;
        }

        internal int GetAvailableExcelRowCountWithoutHeader(string excelTableName)
        {
            return Utility.FindExcelTable(_excelSet, excelTableName).Rows.Count;
        }

        //internal int GetAvailableExcelColumnCount(string excelTableName)
        //{
        //    return Utility.FindExcelTable(_excelSet, excelTableName).Columns.Count;
        //}

        private bool WriteDataRow(DataRow row, dynamic activeSheet, int rowIndexToWrite)
        {
            int realExcelRowIndex = rowIndexToWrite;

            #region Apply Formatter
            string formatterToken = row[Constants.COLUMN_Formatter] as string;
            if (!string.IsNullOrEmpty(formatterToken))
            {
                //IFormatter formatter = row[Constants.COLUMN_Formatter] as IFormatter;
                //string[] formatterTokens = formatterToken.Split(new string[] { FormatterConfig.TokenSplitter }, StringSplitOptions.None);
                string ruleType, rule, formatString;
                FormatterConfig.SplitRuleType_Rule_Formatter(formatterToken, out ruleType, out rule, out formatString);
                IFormatter formatter = FormatterFactory.GetFormatter(ruleType, rule, formatString, this);
                if (formatter != null)
                    formatter.Execute(realExcelRowIndex);
            }
            #endregion

            int excelColIndex = 0;
            int loopColIndex = -1;
            foreach (DataColumn col in row.Table.Columns)
            {
                loopColIndex++;
                if (Utility.IsExtractFromMetadata(col.ColumnName)) //The column with pattern "_..._" is pre-parpared so, should not write to final excel.
                    continue;
                // TODO: Here, if the SheetConfig is null, we use the existing excel data to write directly rather than
                // re-generate.
                if (CurrentSheetConfig == null)
                {
                    activeSheet.Cells[realExcelRowIndex, excelColIndex] = row[col.ColumnName].ToString();
                    continue;
                }
                if (!Utility.IsColumnToOutput(row, loopColIndex))
                    continue;

                excelColIndex++;
                activeSheet.Cells[realExcelRowIndex, excelColIndex] = row[col.ColumnName].ToString();

                string[] columnsAreHyperlink = ((string)row[Constants.COLUMN_HyperlinkColumnIndex]).Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string s in columnsAreHyperlink)
                {
                    if (s.Trim() == loopColIndex.ToString())
                        ExcelOperationWrapper.SetHyperlink(activeSheet, realExcelRowIndex, excelColIndex);
                }
            }

            return true;
        }

        #endregion

        #region Events

        public event EventHandler<GeneralMessageEventArgs> OnGeneralMessageException;

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

    public class GeneralMessageEventArgs : EventArgs
    {
        public string Message { get; private set; }
        public GeneralMessageEventArgs(string message)
        {
            this.Message = message;
        }
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
