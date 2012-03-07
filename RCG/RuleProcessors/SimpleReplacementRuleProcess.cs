using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class SimpleReplacementRuleProcess : BaseRuleProcessor
    {
        private static BaseRuleProcessor _instance = null;
        private static object _lock = new object();

        public override string Process(string source)
        {
            if (string.IsNullOrEmpty(source))
                return source;

            base.PreProcess(source);

            string[] replaceArray = Expressions["exp"].Split(new char[] {','}, StringSplitOptions.RemoveEmptyEntries);
            if (replaceArray != null)
            {
                foreach (string replace in replaceArray)
                {
                    string[] temp = replace.Split(new string[] { "->" }, StringSplitOptions.None);
                    string original = temp[0];
                    string replacement = temp[1];
                    source = source.Replace(original, replacement);
                }
            }
            return source;
        }

        public static BaseRuleProcessor CreateOrGetProcessor(GenProcessor engine)
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new SimpleReplacementRuleProcess(engine);
                    }
                }
            }
            return _instance;
        }

        private SimpleReplacementRuleProcess(GenProcessor engine)
            : base(engine)
        {
        }
    }
}
