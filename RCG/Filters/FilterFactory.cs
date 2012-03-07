using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class FilterFactory
    {
        public static IFilter GetFilter(string filterType, string rule, GenProcessor engine)
        {
            IFilter filter = null;
            switch (filterType)
            {
                case "RegularExpressionFilter":
                    filter = RegularExpressionRuleProcessor.CreateOrGetProcessor(engine) as IFilter;
                    filter.Rule = rule;
                    break;
                default:
                    throw new Exception(string.Format("Rule processor {0} is not recognized", filterType));
            }

            return filter;
        }
    }
}
