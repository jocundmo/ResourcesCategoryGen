using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace RCG
{
    public class RefreshedItemFormatter : BaseFormatter
    {
        public static IFormatter CreateNewProcessor(GenProcessor engine)
        {
            return new RefreshedItemFormatter(engine);
        }

        private RefreshedItemFormatter(GenProcessor engine)
            : base(engine)
        {
            this.Engine = engine;
        }

        #region IFormatter Members

        public override bool Match(DataRow dr, FormatterConfig formatterConfig)
        {
            return Engine.CurrentSheetConfig.Mode == Constants.SHEET_MODE_Refresh;
        }

        #endregion
    }
}
