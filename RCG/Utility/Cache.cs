using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class Cache
    {
        public const string Max_Excel_Row_Count = "_MaxExcelRowCount";
        public const string Max_Exeel_Col_Count = "_MaxExcelColCount";
        public const string Timestamp_DataColumn_Index = "_TimestampDataColumnIndex";
        public const string Primary_DataColumn_Index = "_PrimaryDataColumnIndex";
        public const string Timestamp_ExcelColumn_Index = "_TimestampExcelColumnIndex";
        public const string Primary_ExcelColumn_Index = "_PrimaryExcelColumnIndex";
        public const string DataRow_Exists_Expires_Dict = "_DataRowExistsExpiresDict";

        private static Cache _instance = new Cache();
        public static Cache Instance
        {
            get { return _instance; }
        }

        private Dictionary<string, object> _values = new Dictionary<string, object>();

        public void SetDataRowExistsOrExpiresDictCacheValue(string sheetName, string key, DataRowExistsOrExpires value)
        {
            Dictionary<string, DataRowExistsOrExpires> dict =
                GetCacheValue<Dictionary<string, DataRowExistsOrExpires>>(sheetName, DataRow_Exists_Expires_Dict);

            if (dict == null)
            {
                dict = new Dictionary<string, DataRowExistsOrExpires>();
                SetCacheValue<Dictionary<string, DataRowExistsOrExpires>>(sheetName, DataRow_Exists_Expires_Dict, dict);
            }

            dict[key] = value;
        }

        public DataRowExistsOrExpires GetDataRowExistsOrExpiresDictCacheValue(string sheetName, string key)
        {
            Dictionary<string, DataRowExistsOrExpires> dict = 
                GetCacheValue<Dictionary<string, DataRowExistsOrExpires>>(sheetName, DataRow_Exists_Expires_Dict);

            if (dict == null)
                return DataRowExistsOrExpires.UnKnown;

            if (dict.ContainsKey(key))
                return dict[key];
            else
                return DataRowExistsOrExpires.UnKnown;
        }

        public void SetCacheValue<T>(string sheetName, string key, T value)
        {
            Values[sheetName + key] = value;
        }

        public T GetCacheValue<T>(string sheetName, string key)
        {
            if (Values.ContainsKey(sheetName + key))
                return (T)Values[sheetName + key];
            else
                return default(T);
        }

        public Dictionary<string, object> Values
        {
            get { return _values; }
        }

        private Cache()
        {
            // Nothing to do.
        }
    }
}
