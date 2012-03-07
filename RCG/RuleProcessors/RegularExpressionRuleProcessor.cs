using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace RCG
{
    public class RegularExpressionRuleProcessor : BaseRuleProcessor, IFilter
    {
        private static BaseRuleProcessor _instance = null;
        private static object _lock = new object();

        public override string Process(string source)
        {
            base.PreProcess(source);

            Match m = Regex.Match(source, Expressions["pattern"]);
            if (m.Groups.Count == 1)
                return m.Groups[0].ToString().Trim();
            else
                return m.Groups[m.Groups.Count - 1].ToString().Trim();
        }

        public static BaseRuleProcessor CreateOrGetProcessor(GenProcessor engine)
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new RegularExpressionRuleProcessor(engine);
                    }
                }
            }
            return _instance;
        }

        private RegularExpressionRuleProcessor(GenProcessor engine)
            : base(engine)
        {
        }

        #region IFilter Members

        public bool Match(string source)
        {
            return Regex.IsMatch(source, Rule);
        }

        #endregion
    }
}
