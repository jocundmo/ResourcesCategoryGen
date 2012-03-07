using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Text.RegularExpressions;

namespace RCG
{
    public class DuplicatedItemFormatter : BaseFormatter
    {
        public static IFormatter CreateNewProcessor(GenProcessor engine)
        {
            return new DuplicatedItemFormatter(engine);
        }

        private DuplicatedItemFormatter(GenProcessor engine)
            : base(engine)
        {
            this.Engine = engine;
        }

        #region IFormatter Members

        public override bool Match(DataRow dr, FormatterConfig formatterConfig)
        {
            //bool isAppendUpdateMode = Engine.CurrentSheetConfig.Mode == Constants.SHEET_MODE_Append;
            //bool isMatch = isAppendUpdateMode && IsSmiliarDataRowExistsInExcel(dr, formatterConfig);
            bool isMatch = IsSmiliarDataRowExistsInExcel(dr, formatterConfig);

            return isMatch;
        }

        #endregion

        private bool IsSmiliarDataRowExistsInExcel(DataRow dr, FormatterConfig fc)
        {
            string rowMode = (string)dr[Constants.COLUMN_RowMode];
            if (rowMode == Constants.ROW_MODE_Append)
            {
                string newOne = (string)dr[fc.ExtractFrom];
                if (string.IsNullOrEmpty(newOne))
                    return false;

                foreach (DataRow row in Engine.ExcelSet.Tables[Engine.CurrentSheetConfig.Name].Rows)
                {
                    string oldOne = row[fc.ExtractFrom].ToString();
                    if (string.IsNullOrEmpty(oldOne))
                        continue;
                    if (Regex.IsMatch(newOne, oldOne) ||
                        Regex.IsMatch(oldOne, newOne))
                        return true;
                }
            }
            else if (rowMode == Constants.ROW_MODE_Refresh)
            {
                string newOne = (string)dr[fc.ExtractFrom];
                if (string.IsNullOrEmpty(newOne))
                    return false;

                foreach (DataRow row in Engine.MetadataSet.Tables[Engine.CurrentSheetConfig.Name].Rows)
                {
                    if (row == dr)
                        continue;

                    string oldOne = row[fc.ExtractFrom].ToString();
                    if (string.IsNullOrEmpty(oldOne))
                        continue;
                    if (Regex.IsMatch(newOne, oldOne) ||
                        Regex.IsMatch(oldOne, newOne))
                        return true;

                }
            }
            return false;
        }
    }
}
