using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class FilesTypeRuleProcessor : BaseRuleProcessor
    {
        private static BaseRuleProcessor _instance = null;
        private static object _lock = new object();

        public override string Process(string source)
        {
            base.PreProcess(source);

            string result = source.ToLowerInvariant();
            string excludeValue = Expressions.ContainsKey("exclude")? Expressions["exclude"]:string.Empty;
            string includeValue = Expressions.ContainsKey("include") ? Expressions["include"] : string.Empty;

            if (!string.IsNullOrEmpty(excludeValue))
            {
                foreach (string iv in excludeValue.Split(','))
                {
                    result = result.Replace(iv.ToLowerInvariant(), string.Empty);
                }
            }
            // TODO: implements "include" function.

            return result;
        }

        public static BaseRuleProcessor CreateOrGetProcessor(GenProcessor engine)
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new FilesTypeRuleProcessor(engine);
                    }
                }
            }
            return _instance;
        }

        private FilesTypeRuleProcessor(GenProcessor engine)
            : base(engine)
        {
        }
    }
}
