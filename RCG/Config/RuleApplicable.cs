using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public abstract class RuleApplicable
    {
        public string ExtractFrom { get; set; }
        public string RuleType { get; set; }
        public string Rule { get; set; }

        public RuleApplicable(string extractFrom, string ruleType, string rule)
        {
            this.ExtractFrom = extractFrom;
            this.RuleType = ruleType;
            this.Rule = rule;
        }
    }
}
