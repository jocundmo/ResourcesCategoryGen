using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Data;

namespace RCG
{
    public class NumberCompareFormatter : BaseFormatter
    {
        public static IFormatter CreateNewProcessor(GenProcessor engine)
        {
            return new NumberCompareFormatter(engine);
        }

        private NumberCompareFormatter(GenProcessor engine)
            : base(engine)
        {
            this.Engine = engine;
        }

        #region IFormatter Members

        public override bool Match(DataRow dr, FormatterConfig formatterConfig)
        {
            string source = Utility.GetDataRowContent(dr, formatterConfig.ExtractFrom);
            string oper = Rule.Split(':')[0];
            string num = Rule.Split(':')[1];
            long n = long.Parse(num);

            if (oper.Trim() == "less_than")
                return long.Parse(source) < n;
            else if (oper.Trim() == "greater_than")
                return long.Parse(source) > n;
            else if (oper.Trim() == "equal")
                return long.Parse(source) == n;

            throw new ArgumentException(string.Format("Not recognized operator {0} ...", oper));
        }

        #endregion
    }
}
