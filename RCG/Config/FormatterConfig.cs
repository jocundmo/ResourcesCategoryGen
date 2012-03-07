using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class FormatterConfig : RuleApplicable
    {
        public string Name { get; set; }
        public bool Enabled { get; set; }
        public string FormatString { get; set; }

        public FormatterConfig(string name, string extractFrom, string ruleType, string rule, bool enabled, string formatString)
            : base(extractFrom, ruleType, rule)
        {
            this.Name = name;
            this.Enabled = enabled;
            this.FormatString = formatString;
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
