using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class FormatterConfig : RuleApplicable
    {
        public const string TokenSplitter = "`";

        public string Name { get; set; }
        public bool Enabled { get; set; }
        public string FormatString { get; set; }
        public string Token { get; private set; } // Token is representing the unique formatter when user factory go get.

        public FormatterConfig(string name, string extractFrom, string ruleType, string rule, bool enabled, string formatString)
            : base(extractFrom, ruleType, rule)
        {
            this.Name = name;
            this.Enabled = enabled;
            this.FormatString = formatString;
            this.Token = this.RuleType + TokenSplitter + this.Rule + TokenSplitter + this.FormatString;
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
