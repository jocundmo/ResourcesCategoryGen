using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Collections.ObjectModel;

namespace RCG
{
    public abstract class BaseRuleProcessor
    {
        protected const string ParameterSource = "{source}";
        protected const string ParameterNow = "{now}";
        protected const string ParameterEvalStatementEnd = "{end}";

        private Dictionary<string, string> _expressions = new Dictionary<string, string>();

        public string Rule { get; set; }
        public GenProcessor Processor { get; private set; }
        public bool Stateless { get; protected set; }
        public bool Parsed { get; private set; }

        public Dictionary<string, string> Expressions
        {
            get { return _expressions; }
        }

        protected BaseRuleProcessor(GenProcessor engine)
        {
            this.Processor = engine;
        }

        protected virtual void Parse()
        {
            Expressions.Clear();
            string[] statements = Rule.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            if (statements != null)
            {
                foreach (string statement in statements)
                {
                    string[] tempArray = statement.Split(new string[] {":="}, StringSplitOptions.None);
                    if (tempArray == null || tempArray.Length < 2)
                        throw new Exception(string.Format("The rule {0} is not recognized", Rule));
                    Expressions[tempArray[0]] = tempArray[1];
                }
            }
            Parsed = true;
        }

        protected virtual void PreProcess(string source)
        {
            Parse();
            ReplaceInternalSystemParameters(source);
        }

        private void ReplaceInternalSystemParameters(string source)
        {
            Dictionary<string, string> replacements = new Dictionary<string, string>();

            foreach (KeyValuePair<string, string> kvp in Expressions)
            {
                replacements[kvp.Key] = kvp.Value;

                replacements[kvp.Key] = VariableRefresher.RefreshSystemVariable(replacements[kvp.Key]);
                replacements[kvp.Key] = VariableRefresher.RefreshRuleProcessorVariable(replacements[kvp.Key], source);
                //if (kvp.Value.Contains(ParameterSource))
                //    replacements[kvp.Key] = replacements[kvp.Key].Replace(ParameterSource, source);

                //if (kvp.Value.Contains(ParameterNow))
                //    replacements[kvp.Key] = replacements[kvp.Key].Replace(ParameterNow, DateTime.Now.ToString());

                //if (kvp.Value.Contains(ParameterEvalStatementEnd))
                //    replacements[kvp.Key] = replacements[kvp.Key].Replace(ParameterEvalStatementEnd, ";");
            }

            foreach (KeyValuePair<string, string> kvp in replacements)
            {
                Expressions[kvp.Key] = kvp.Value;
            }
        }

        public abstract string Process(string source);
    }
}
