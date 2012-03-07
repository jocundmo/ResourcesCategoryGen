using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class ConditionalRuleProcessor : BaseRuleProcessor
    {
        private static BaseRuleProcessor _instance = null;
        private static object _lock = new object();

        public static BaseRuleProcessor CreateOrGetProcessor(GenProcessor engine)
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new ConditionalRuleProcessor(engine);
                    }
                }
            }
            return _instance;
        }

        private ConditionalRuleProcessor(GenProcessor engine)
            : base(engine)
        {
        }


        public override string Process(string source)
        {
            base.PreProcess(source);

            string[] statements = Expressions["statement"].Split(',');
            string evalResult = string.Empty;
            foreach (string statement in statements)
            {
                evalResult = Evaluator.Eval(statement).ToString();
                if (!string.IsNullOrEmpty(evalResult))
                    break;
            }
            return evalResult;
        }
    }
}
