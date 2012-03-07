using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class VariableRefresher
    {
        protected const string ParameterNow = "{now}";

        protected const string ParameterSource = "{source}";
        protected const string ParameterEvalStatementEnd = "{end}";

        public static string RefreshRuleProcessorVariable(string originalValue, string source)
        {
            if (originalValue.Contains(ParameterSource))
                originalValue = originalValue.Replace(ParameterSource, source);

            if (originalValue.Contains(ParameterEvalStatementEnd))
                originalValue = originalValue.Replace(ParameterEvalStatementEnd, ";");

            return originalValue;
        }

        public static string RefreshSystemVariable(string originalValue)
        {
            return RefreshSystemVariable(originalValue, string.Empty);
        }

        public static string RefreshSystemVariable(string originalValue, string DateTimeFormat)
        {
            if (originalValue.Contains(ParameterNow))
            {
                if (string.IsNullOrEmpty(DateTimeFormat))
                    originalValue = originalValue.Replace(ParameterNow, DateTime.Now.ToString());
                else
                    originalValue = originalValue.Replace(ParameterNow, DateTime.Now.ToString(DateTimeFormat));
            }

            return originalValue;
        }
    }
}
