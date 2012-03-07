using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class FilterConfig : RuleApplicable
    {
        public string Name { get; set; }
        public bool Enabled { get; set; }

        public FilterConfig(string name, string extractFrom, string ruleType, string rule, bool enabled)
            : base(extractFrom, ruleType, rule)
        {
            this.Name = name;
            this.Enabled = enabled;
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
