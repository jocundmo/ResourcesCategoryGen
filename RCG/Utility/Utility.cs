using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO;

namespace RCG
{
    public static class Utility
    {
        public static void ParseFileTypeConfig(string fileTypeConfig, out string fileType, out SearchOption searchOption)
        {
            string[] r = fileTypeConfig.Split(new char[] {'(', ')'}, StringSplitOptions.RemoveEmptyEntries);
            fileType = r[0];
            if (r.Length > 1 && r[1] == ("/s"))
                searchOption = SearchOption.AllDirectories;
            else
                searchOption = SearchOption.TopDirectoryOnly;
        }

        public static string ConvertVolumeLabelPathToPhysicalPath(string volumeLabelPath)
        {
            string volumeLabel = Regex.Match(volumeLabelPath, @"^[\w\d ]+(?=\\)").ToString();

            foreach (string ld in Directory.GetLogicalDrives())
            {
                DriveInfo di = new DriveInfo(ld);
                if (di.IsReady && di.VolumeLabel == volumeLabel)
                {
                    string path = Regex.Replace(volumeLabelPath, @"^[\w\d ]+\\", ld);
                    return path;
                }
            }
            //throw new Exception(string.Format("Volume label {0} is not recognized...", volumeLabelPath));
            return string.Empty;
        }

        public static LocationType GetLocationType(string path)
        {
            LocationType r = LocationType.Unknown;

            if (Regex.IsMatch(path, @"^[A-Za-z]:\\")) // e.g.  c:\
                r = LocationType.Physical;
            else if (Regex.IsMatch(path, @"^\\\\\d{1,3}?\.\d{1,3}\.\d{1,3}\.\d{1,3}\\")) // e.g. \\192.168.2.165\
                r = LocationType.Network;
            else if (Regex.IsMatch(path, @"\\\\[A-Za-z]+\\")) // e.g. \\rabook\
                r = LocationType.Network;
            else if (Regex.IsMatch(path, @"^[\w\d ]+\\")) // e.g. 808G_01\
                r = LocationType.VolumeLabel;

            if (r == LocationType.Unknown)
                throw new Exception(string.Format("Input path {0} is not recognized...", path));

            return r;
        }

        public static void CalFolderSize(ref long size, ref int fileCount, ref int subFolderCount, ref string filesType, DirectoryInfo di)
        {
            //FileInfo[] fileList = di.GetFiles();
            FileInfo[] fileList = null;

            if (di.Exists)
            {
                fileList = di.GetFiles();
            }
            else
            {
                fileList = new FileInfo[1];
                fileList[0] = new FileInfo(di.FullName);
            }
            fileCount += fileList.Length;
            foreach (FileInfo f in fileList)
            {
                string extension = Path.GetExtension(f.FullName);
                if (!filesType.Contains(extension))
                    filesType += extension;
                size += f.Length;
            }
            if (di.Exists)
            {
                DirectoryInfo[] subFolderList = di.GetDirectories();
                if (subFolderList.Length != 0)
                {
                    subFolderCount += subFolderList.Length;
                    foreach (DirectoryInfo innerDi in subFolderList)
                    {
                        CalFolderSize(ref size, ref fileCount, ref subFolderCount, ref filesType, innerDi);
                    }
                }
            }
        }

        public static DataTable FindMetadataTable(DataSet metadataSet, string metadataTableName)
        {
            foreach (DataTable dt in metadataSet.Tables)
            {
                if (dt.TableName == metadataTableName)
                    return dt;
            }
            return null;
            //throw new Exception(string.Format("Metadata table {0} not recognized...", metadataTableName));
        }

        public static DataTable FindExcelTable(DataSet excelSet, string excelTableName)
        {
            foreach (DataTable dt in excelSet.Tables)
            {
                if (dt.TableName == excelTableName)
                    return dt;
            }
            return null;
            //throw new Exception(string.Format("Excel table {0} not recognized...", excelTableName));
        }

        public static bool IsExtractFromMetadata(string path)
        {
            return Regex.IsMatch(path, "_.*?_");
        }

        public static bool IsValidExtractFrom(string path)
        {
            bool valid = true;
            valid = (path != Constants.PREDEFINED_AutoIncrease) &&
                !string.IsNullOrEmpty(path);
            return valid;
        }


        public static bool IsPredefinedColumn(string path)
        {
            return path.Trim() == Constants.PREDEFINED_AutoIncrease;
        }

        public static bool IsColumnToOutput(DataRow row,  int currentColumnIndex)
        {
            string[] columnsCouldOutput = ((string)row[Constants.COLUMN_OutputColumnIndex]).Split(new char[] {','}, StringSplitOptions.RemoveEmptyEntries);
            foreach (string s in columnsCouldOutput)
            {
                if (s.Trim() == currentColumnIndex.ToString())
                    return true;
            }
            return false;
            //ColumnConfig ccToFound = null;
            //foreach (ColumnConfig cc in currentSheetConfig.Columns)
            //{
            //    if (cc.Enabled && cc.Name == path)
            //    {
            //        ccToFound = cc;
            //        break;
            //    }
            //}
            //if (ccToFound == null)
            //    throw new Exception(string.Format("Column {0} not found or not enabled", path));
            //return ccToFound.Output;
        }

        public static string GetDataRowContent(DataRow row, string extractFrom)
        {
            if (IsValidExtractFrom(extractFrom.Trim()))
                return row[extractFrom].ToString();
            else
                return string.Empty;
        }

        public static string GetAttributeValue(XmlNode node, string key)
        {
            return GetAttributeValue(node, key, string.Empty);
        }

        public static string GetAttributeValue(XmlNode node, string key, string defaultValue)
        {
            XmlAttribute attribute = null;
            foreach (XmlAttribute attr in node.Attributes)
            {
                if (attr.Name.Equals(key, StringComparison.OrdinalIgnoreCase))
                {
                    attribute = attr;
                    break;
                }
            }
            //XmlAttribute attribute = node.Attributes[key];
            string result = defaultValue;

            if (attribute != null)
                result = attribute.Value.Trim();

            return VariableRefresher.RefreshSystemVariable(result, "yyyyMMdd-hhmmss");
            //else
            //    return defaultValue;
        }

    }
}
