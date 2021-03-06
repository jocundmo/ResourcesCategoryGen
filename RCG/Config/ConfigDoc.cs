﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Text.RegularExpressions;

namespace RCG
{
    public class ConfigDoc
    {
        public static bool ConvertToBoolean(string b)
        {
            if (0 == string.Compare(b, bool.FalseString, true))
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        public static int ConvertToInt(string b)
        {
            return int.Parse(b);
        }

        private List<SheetConfig> _sheets = new List<SheetConfig>();
        public List<SheetConfig> Sheets { get { return _sheets; } }
        public string TemplatePath { get; set; }
        public string BaselinePath { get; set; }
        public string OutputPath { get; set; }

        public void Read(string configFileName)
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(configFileName);
            XmlNodeList xnlRoot = xml.SelectNodes("//Document");
            if (xnlRoot.Count != 1)
                throw new Exception("There should be one (and only one) <document> section defined");

            this.TemplatePath = Utility.GetAttributeValue(xnlRoot[0], "templatePath", string.Empty);
            this.BaselinePath = Utility.GetAttributeValue(xnlRoot[0], "baselinePath", string.Empty);
            this.OutputPath = Utility.GetAttributeValue(xnlRoot[0], "outputPath", DateTime.Now.ToString("yyyyMMdd"));

            XmlNodeList xnlSheets = ((XmlElement)xnlRoot[0]).SelectNodes("Sheets/Sheet");
            foreach (XmlElement xeSheet in xnlSheets)
            {
                string sheetName = Utility.GetAttributeValue(xeSheet, "name");
                string sheetMode = Utility.GetAttributeValue(xeSheet, "mode", "refresh");
                int sheetMaxRowCount = ConfigDoc.ConvertToInt(Utility.GetAttributeValue(xeSheet, "maxRowCount", "3000"));
                bool sheetEnabled = ConfigDoc.ConvertToBoolean(Utility.GetAttributeValue(xeSheet, "enabled", "true"));
                SheetConfig sheet = new SheetConfig(sheetName, sheetEnabled, sheetMode, sheetMaxRowCount);
                // Read [Columns > Column]
                XmlNodeList xnlColumns = xeSheet.SelectNodes("Columns/Column");
                foreach (XmlElement xeColumn in xnlColumns)
                {
                    string columnName = Utility.GetAttributeValue(xeColumn, "name");

                    string ruleType = Utility.GetAttributeValue(xeColumn, "ruleType");
                    string rule = Utility.GetAttributeValue(xeColumn, "rule");
                    string extractFrom = Utility.GetAttributeValue(xeColumn, "extractFrom");

                    bool columnEnabled = ConfigDoc.ConvertToBoolean(Utility.GetAttributeValue(xeColumn, "enabled", "true"));
                    bool columnPrimary = ConfigDoc.ConvertToBoolean(Utility.GetAttributeValue(xeColumn, "primary", "false"));
                    bool columnTimestamp = ConfigDoc.ConvertToBoolean(Utility.GetAttributeValue(xeColumn, "timestamp", "false"));
                    bool columnOutput = ConfigDoc.ConvertToBoolean(Utility.GetAttributeValue(xeColumn, "output", "true"));
                    sheet.Columns.Add(new ColumnConfig(columnName, columnName, extractFrom, ruleType, rule, columnEnabled, columnPrimary, columnTimestamp, columnOutput));
                }

                // Raad [Filters > Filter]
                XmlNodeList xnlFilters = xeSheet.SelectNodes("Filters/Filter");
                foreach (XmlElement xeFilter in xnlFilters)
                {
                    string filterName = Utility.GetAttributeValue(xeFilter, "name");

                    string ruleType = Utility.GetAttributeValue(xeFilter, "ruleType");
                    string rule = Utility.GetAttributeValue(xeFilter, "rule");
                    string extractFrom = Utility.GetAttributeValue(xeFilter, "extractFrom");

                    bool filterEnabled = ConfigDoc.ConvertToBoolean(Utility.GetAttributeValue(xeFilter, "enabled", "true"));
                    sheet.Filters.Add(new FilterConfig(filterName, extractFrom, ruleType, rule, filterEnabled));
                }

                // Raad [Formatters > Formatter]
                XmlNodeList xnlFormatters = xeSheet.SelectNodes("Formatters/Formatter");
                foreach (XmlElement xeFormatter in xnlFormatters)
                {
                    string formatterName = Utility.GetAttributeValue(xeFormatter, "name");

                    string ruleType = Utility.GetAttributeValue(xeFormatter, "ruleType");
                    string rule = Utility.GetAttributeValue(xeFormatter, "rule");
                    string extractFrom = Utility.GetAttributeValue(xeFormatter, "extractFrom");

                    string formatString = Utility.GetAttributeValue(xeFormatter, "formatString");
                    bool formatterEnabled = ConfigDoc.ConvertToBoolean(Utility.GetAttributeValue(xeFormatter, "enabled", "true"));
                    sheet.Formatters.Add(new FormatterConfig(formatterName, extractFrom, ruleType, rule, formatterEnabled, formatString));
                }

                // Read [Locations > Location]
                LocationConfig locationConfig = null;
                XmlNodeList xnlLocations = xeSheet.SelectNodes("Locations/Location");
                foreach (XmlElement xeLocation in xnlLocations)
                {
                    locationConfig = new LocationConfig();
                    locationConfig.Path = Utility.GetAttributeValue(xeLocation, "path");
                    string include = Utility.GetAttributeValue(xeLocation, "include", "folder");
                    locationConfig.IncludeFolder = (include == "Folder" || include == "folder");
                    if (!locationConfig.IncludeFolder)
                    {
                        string[] fileTypes = include.Split(new char[] {';'}, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string ft in fileTypes)
                        {
                            string fileType = string.Empty;
                            System.IO.SearchOption searchOption = System.IO.SearchOption.TopDirectoryOnly;
                            Utility.ParseFileTypeConfig(ft, out fileType, out searchOption);
                            locationConfig.IncludeFileTypes.Add(new FileTypeConfig(fileType, searchOption));                           
                        }
                    }
                    locationConfig.Enabled = ConfigDoc.ConvertToBoolean(Utility.GetAttributeValue(xeLocation, "enabled", "true"));

                    sheet.Locations.Add(locationConfig);
                }

                this.Sheets.Add(sheet);
            }
        }

        public void Validate()
        {
            foreach (SheetConfig sheetConfig in _sheets)
            {
                bool foundPrimaryColumn = false;
                bool foundTimestampColumn = false;
                foreach (ColumnConfig columnConfig in sheetConfig.Columns)
                {
                    if (columnConfig.Primary)
                    {
                        if (foundPrimaryColumn)
                            throw new Exception(string.Format("Sheet {0} only can make one Primary column.", sheetConfig.Name));
                        foundPrimaryColumn = true;
                    }
                    if (columnConfig.Timestamp)
                    {
                        if (foundTimestampColumn)
                            throw new Exception(string.Format("Sheet {0} only can make one Timestamp column.", sheetConfig.Name));
                        foundTimestampColumn = true;
                    }
                }
                if (!foundPrimaryColumn)
                    throw new Exception(string.Format("Sheet {0} no Primary column found, make sure one is defined", sheetConfig.Name));
                if (!foundTimestampColumn)
                    throw new Exception(string.Format("Sheet {0} no Timestamp column found, make sure one is defined", sheetConfig.Name));
            }
        }
    }
}
