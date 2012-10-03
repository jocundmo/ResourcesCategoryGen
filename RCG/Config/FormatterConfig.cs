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

        public static void SplitRuleType_Rule_Formatter(string token, out string ruleType, out string rule, out string formatString)
        {
            string[] formatterTokens = token.Split(new string[] { FormatterConfig.TokenSplitter }, StringSplitOptions.None);
            if (formatterTokens.Length > 2)
            {
                ruleType = formatterTokens[0];
                rule = formatterTokens[1];
                formatString = formatterTokens[2];
            }
            else
            {
                ruleType = string.Empty;
                rule = string.Empty;
                formatString = string.Empty;
            }
        }

        public static FormatterConfig GetDefaultAppendedItemFormatterConfig()
        {
            FormatterConfig appendedItemFormatterCofig = new FormatterConfig("defaultAppendedItem_formatter", "", "AppendedItemFormatter", Constants.FORMATTER_Internal_AppendedItem, true, "font-bold");
            return appendedItemFormatterCofig;
        }

        public static FormatterConfig GetDefaultDeletedItemFormatterConfig()
        {
            FormatterConfig deletedItemFormatterCofig = new FormatterConfig("defaultDeletedItem_formatter", "", "DeletedItemFormatter", Constants.FORMATTER_Internal_DeletedItem, true, "font-strikethrough");
            return deletedItemFormatterCofig;
        }

        internal static FormatterConfig GetDefaultUpdatedItemFormatterConfig()
        {
            FormatterConfig updatedItemFormatterCofig = new FormatterConfig("defaultDeletedItem_formatter", "", "DeletedItemFormatter", Constants.FORMATTER_Internal_UpdatedItem, true, "font-italic");
            return updatedItemFormatterCofig;
        }
    }
}
