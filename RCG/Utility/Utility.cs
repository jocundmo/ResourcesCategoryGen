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
        public static void MergeXml(XmlDocument baselineDoc, XmlDocument snippetDoc, string commonNodeXPathExpression, bool onlyAttribute)
        {
            XmlElement xeOri = (XmlElement)baselineDoc.SelectSingleNode(commonNodeXPathExpression);
            XmlElement xeNew = (XmlElement)snippetDoc.SelectSingleNode(commonNodeXPathExpression);

            foreach (XmlAttribute attrNew in xeNew.Attributes)
            {
                XmlAttribute findAttri = null;
                foreach (XmlAttribute attrToBeReplaced in xeOri.Attributes)
                {
                    if (attrToBeReplaced.Name.Equals(attrNew.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        findAttri = attrToBeReplaced;
                        break;
                    }
                }
                if (findAttri != null)
                {
                    findAttri.Value = attrNew.Value;
                }
            }
            if (!onlyAttribute)
                xeOri.InnerXml = xeNew.InnerXml;
        }

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

        public static LocationType GetLocationType(string path, ref string value)
        {
            LocationType r = LocationType.Unknown;

            Match m = null;
            m = Regex.Match(path, @"^[A-Za-z]:\\");
            if (m.Groups.Count > 0 && m.Groups[0].Value != string.Empty)
            {
                value = m.Groups[0].Value;
                return LocationType.Physical;
            }
            m = Regex.Match(path, @"^\\\\\d{1,3}?\.\d{1,3}\.\d{1,3}\.\d{1,3}\\");
            if (m.Groups.Count > 0 && m.Groups[0].Value != string.Empty)
            {
                m = Regex.Match(path, @"^\\\\\d{1,3}?\.\d{1,3}\.\d{1,3}\.\d{1,3}\\([A-Za-z0-9_-]+)\\");
                if (m.Groups.Count == 1)
                    value = m.Groups[0].ToString().Trim();
                else
                    value = m.Groups[m.Groups.Count - 1].ToString().Trim();
                //value = m.Groups[0].Value;
                return LocationType.Network;
            }
            m = Regex.Match(path, @"^\\\\[A-Za-z0-9]+\\");
            if (m.Groups.Count > 0 && m.Groups[0].Value != string.Empty)
            {
                m = Regex.Match(path, @"^\\\\[A-Za-z0-9]+\\([A-Za-z0-9_-]+)\\");
                if (m.Groups.Count == 1)
                    value = m.Groups[0].ToString().Trim();
                else
                    value = m.Groups[m.Groups.Count - 1].ToString().Trim();
                return LocationType.Network;
            }
            m = Regex.Match(path, @"^[\w\d ]+\\");
            if (m.Groups.Count > 0 && m.Groups[0].Value != string.Empty)
            {
                value = m.Groups[0].Value;
                return LocationType.VolumeLabel;
            }
            //if (Regex.IsMatch(path, @"^[A-Za-z]:\\")) // e.g.  c:\
            //    r = LocationType.Physical;
            //else if (Regex.IsMatch(path, @"^\\\\\d{1,3}?\.\d{1,3}\.\d{1,3}\.\d{1,3}\\")) // e.g. \\192.168.2.165\
            //    r = LocationType.Network;
            //else if (Regex.IsMatch(path, @"\\\\[A-Za-z0-9]+\\")) // e.g. \\rabook\
            //    r = LocationType.Network;
            //else if (Regex.IsMatch(path, @"^[\w\d ]+\\")) // e.g. 808G_01\
            //    r = LocationType.VolumeLabel;

            if (r == LocationType.Unknown)
                throw new Exception(string.Format("Input path {0} is not recognized...", path));

            return r;
        }
        public static LocationType GetLocationType(string path)
        {
            string e = string.Empty;
            return GetLocationType(path, ref e);
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


        internal static DataRow FindSameRow(DataRow row, DataTable tableExcel)
        {
            foreach (DataRow dr in tableExcel.Rows)
            {
                string primaryKey = (string)dr[(int)dr[Constants.COLUMN_PrimaryColumnIndex]];
                if (primaryKey == (string)row[(int)row[Constants.COLUMN_PrimaryColumnIndex]])
                {
                    return dr;
                }
            }
            return null;
        }

        public static int MoreCharToInt(string value)
        {
            int rtn = 0;
            int powIndex = 0;

            for (int i = value.Length - 1; i >= 0; i--)
            {
                int tmpInt = value[i];
                tmpInt -= 64;

                rtn += (int)Math.Pow(26, powIndex) * tmpInt;
                powIndex++;
            }

            return rtn;
        }

        public static string IntToMoreChar(int value)
        {
            string rtn = string.Empty;
            List<int> iList = new List<int>();

            //To single Int
            while (value / 26 != 0 || value % 26 != 0)
            {
                iList.Add(value % 26);
                value /= 26;
            }

            //Change 0 To 26
            for (int j = 0; j < iList.Count - 1; j++)
            {
                if (iList[j] == 0)
                {
                    iList[j + 1] -= 1;
                    iList[j] = 26;
                }
            }

            //Remove 0 at last
            if (iList[iList.Count - 1] == 0)
            {
                iList.Remove(iList[iList.Count - 1]);
            }

            //To String
            for (int j = iList.Count - 1; j >= 0; j--)
            {
                char c = (char)(iList[j] + 64);
                rtn += c.ToString();
            }

            return rtn;
        }

        internal static bool IsLocationAccessable(string path)
        {
            try
            {
                return Directory.Exists(Utility.ConvertVolumeLabelPathToPhysicalPath(path));
            }
            catch (IOException)
            {
                return false;
            }
        }
    }
}
