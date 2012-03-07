using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace RCG
{
    public class AppendedItemFormatter : BaseFormatter
    {
        public static IFormatter CreateNewProcessor(GenProcessor engine)
        {
            return new AppendedItemFormatter(engine);
        }

        private AppendedItemFormatter(GenProcessor engine)
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
            bool isDataRowExistsInExcel = (Engine.IsDataRowExistsOrExpires(dr) != DataRowExistsOrExpires.NotExists);
            bool isMatch = isAppendUpdateMode && !isDataRowExistsInExcel;

            return isMatch;
        }

        #endregion
    }
}
