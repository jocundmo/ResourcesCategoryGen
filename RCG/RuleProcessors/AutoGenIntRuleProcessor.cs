using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class AutoGenIntRuleProcessor : BaseRuleProcessor
    {
        private string currentSheetName = string.Empty;
        private static BaseRuleProcessor _instance = null;
        private static object _lock = new object();
        private int seed = 1;

        public override string Process(string source)
        {
            if (string.IsNullOrEmpty(currentSheetName))
                currentSheetName = Processor.CurrentSheetConfig.Name;
            if (currentSheetName != Processor.CurrentSheetConfig.Name)
            {
                currentSheetName = Processor.CurrentSheetConfig.Name;
                seed = 1;
            }
            return (seed ++).ToString();
        }

        public static BaseRuleProcessor CreateOrGetProcessor(GenProcessor engine)
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = new AutoGenIntRuleProcessor(engine);
                    }
                }
            }
            return _instance;
        }

        private AutoGenIntRuleProcessor(GenProcessor engine)
            : base(engine)
        {
            this.Stateless = false;
        }
    }
}
