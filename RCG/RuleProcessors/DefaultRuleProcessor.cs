using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class DefaultRuleProcessor : BaseRuleProcessor
    {
        private static BaseRuleProcessor _instance = null;
        private static object _lock = new object();

        public override string Process(string source)
        {
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
                        _instance = new DefaultRuleProcessor(engine);
                    }
                }
            }
            return _instance;
        }

        private DefaultRuleProcessor(GenProcessor engine)
            : base(engine)
        {
        }
    }
}
