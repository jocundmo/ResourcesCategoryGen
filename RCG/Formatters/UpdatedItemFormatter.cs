using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace RCG
{
    public class UpdatedItemFormatter : BaseFormatter
    {
        public static IFormatter CreateNewProcessor(GenProcessor engine)
        {
            return new UpdatedItemFormatter(engine);
        }

        private UpdatedItemFormatter(GenProcessor engine)
            : base(engine)
        {
            this.Engine = engine;
        }

        #region IFormatter Members

        public override bool Match(DataRow dr, FormatterConfig formatterConfig)
        {
            if (Engine.CurrentSheetConfig.Mode == Constants.SHEET_MODE_Refresh)
                return false;

            bool isAppendUpdateMode = Engine.CurrentSheetConfig.Mode == Constants.SHEET_MODE_Append;
            bool isDataRowExpiresInExcel = (Engine.IsDataRowExistsOrExpires(dr) == DataRowExistsOrExpires.ExistsAndExpires);
            bool isMatch = isAppendUpdateMode && isDataRowExpiresInExcel;

            return isMatch;
        }

        #endregion
    }
}
