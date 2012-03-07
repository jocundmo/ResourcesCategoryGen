using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class ColumnConfig : RuleApplicable
    {
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public bool Primary { get; set; }
        public bool Output { get; set; }
        public bool Timestamp { get; set; }
        public bool Enabled { get; set; }

        public ColumnConfig(string name, string extractFrom, string ruleType, string rule)
            : this(name, extractFrom, name, ruleType, rule, true, false, false, true)
        {
            // Nothing to do.
        }

        public ColumnConfig(string name, string displayName, string extractFrom, string ruleType, string rule, bool enabled, bool primary, bool timestamp, bool output)
            : base(extractFrom, ruleType, rule)
        {
            this.Name = name;
            this.DisplayName = displayName;
            this.Enabled = enabled;
            this.Primary = primary;
            this.Timestamp = timestamp;
            this.Output = output;
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
