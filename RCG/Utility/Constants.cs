using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public static class Constants
    {
        public const int INT_NOT_FOUND_INDEX = -1;

        public const string PREDEFINED_AutoIncrease = "_AutoIncrease_";
        public const string COLUMN_Path = "_Path_";
        public const string COLUMN_LastModified = "_LastModified_";
        public const string COLUMN_Size = "_Size_";
        public const string COLUMN_Attributes = "_Attributes_";
        public const string COLUMN_FileCount = "_FileCount_";
        public const string COLUMN_SubFolderCount = "_SubFolderCount_";
        public const string COLUMN_FromFile = "_FromFile_";
        public const string COLUMN_FilesType = "_FilesType_";
        public const string COLUMN_FilterFlag = "_*FilterFlag*_";
        public const string COLUMN_RowMode = "_*RowMode*_";
        public const string COLUMN_Formatter = "_*FormatString*_";
        public const string COLUMN_Tag = "_*Tag*_";
        public const string COLUMN_PrimaryColumnIndex = "_*PrimaryColumnIndex*_";
        public const string COLUMN_TimestampColumnIndex = "_*TimestampColumnIndex*_";
        public const string COLUMN_OutputColumnIndex = "_*OutputColumnIndex*_";
        public const string COLUMN_AutoIncreaseColumnIndex = "_*AutoIncreaseColumnIndex*_";
        public const string COLUMN_LocationFrom = "_*LocationFrom*_";

        public const int HEADER_ROW_INDEX = 1;

        public const string SHEET_MODE_Refresh = "refresh";
        public const string SHEET_MODE_Append = "append";

        public const string ROW_MODE_Filtered = "filtered";
        public const string ROW_MODE_Ignored = "ignored";
        public const string ROW_MODE_Append = "append";
        public const string ROW_MODE_Update = "update";
        public const string ROW_MODE_Refresh = "refresh";
        public const string ROW_MODE_Deleted = "deleted";

        public const string FORMATTER_Internal_AppendedItem = "_*static.appended*_";
        public const string FORMATTER_Internal_DeletedItem = "_*static.deleted*_";
        public const string FORMATTER_Internal_UpdatedItem = "_*static.updated*_";
    }

    public enum LocationType
    {
        Unknown = 0,
        Physical = 1,
        Network = 2,
        VolumeLabel = 3
    }

    public enum DataRowExistsOrExpires
    {
        UnKnown = 0,
        NotExists = 1,
        ExistsButNotExpires = 2,
        ExistsAndExpires = 3
    }
}
