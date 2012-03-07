using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Data;

namespace RCG
{
    public class DateTimeCompareFormatter : BaseFormatter
    {
        public static IFormatter CreateNewProcessor(GenProcessor engine)
        {
            return new DateTimeCompareFormatter(engine);
        }

        private DateTimeCompareFormatter(GenProcessor engine)
            : base(engine)
        {
            this.Engine = engine;
        }

        #region IFormatter Members

        public override bool Match(DataRow dr, FormatterConfig formatterConfig)
        {
            string source = Utility.GetDataRowContent(dr, formatterConfig.ExtractFrom);
            string oper = Rule.Split(':')[0];
            string datetime = Rule.Split(':')[1];
            DateTime n = DateTime.MinValue;

            if (datetime.Trim() == "{%now%}")
                n = DateTime.Now;
            else
                n = DateTime.Parse(datetime);

            if (oper.Trim() == "less_than")
                return DateTime.Parse(source) < n;
            else if (oper.Trim() == "greater_than")
                return DateTime.Parse(source) > n;
            else if (oper.Trim() == "equal")
                return DateTime.Parse(source) == n;

            throw new ArgumentException(string.Format("Not recognized operator {0} ...", oper));
        }

        #endregion
    }
}
