using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class DeletedItemFormatter : BaseFormatter
    {
        public static IFormatter CreateNewProcessor(GenProcessor engine)
        {
            return new DeletedItemFormatter(engine);
        }

        private DeletedItemFormatter(GenProcessor engine)
            : base(engine)
        {
            this.Engine = engine;
        }

        public override bool Match(System.Data.DataRow dr, FormatterConfig formatterConfig)
        {
            if (Engine.CurrentSheetConfig.Mode == Constants.SHEET_MODE_Refresh)
                return false;

            bool isMatch = ((string)dr[Constants.COLUMN_RowMode] == Constants.ROW_MODE_Deleted);
            return isMatch;
        }
    }
}
