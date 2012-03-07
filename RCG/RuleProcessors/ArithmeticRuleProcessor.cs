using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class ArithmeticRuleProcessor : BaseRuleProcessor
    {
        private static BaseRuleProcessor _instance = null;
        private static object _lock = new object();

        public override string Process(string source)
        {
            base.PreProcess(source);

            string statement = Expressions["exp"];
            double result = double.Parse((string)Evaluator.Eval(statement));

            return result.ToString(Expressions["format"]);
        }

        public static BaseRuleProcessor CreateOrGetProcessor(GenProcessor engine)
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new ArithmeticRuleProcessor(engine);
                    }
                }
            }
            return _instance;
        }

        private ArithmeticRuleProcessor(GenProcessor engine)
            : base(engine)
        {
        }
    }
}
