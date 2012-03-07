using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class RuleProcessorFactory
    {
        public static BaseRuleProcessor GetRuleProcessor(ColumnConfig columnConfig, GenProcessor engine)
        {
            BaseRuleProcessor processor = null;
            switch (columnConfig.RuleType)
            {
                case "RegularExpressionRuleProcessor":
                    processor = RegularExpressionRuleProcessor.CreateOrGetProcessor(engine);
                    processor.Rule = columnConfig.Rule;
                    break;
                case "AutoGenIntRuleProcessor":
                    processor = AutoGenIntRuleProcessor.CreateOrGetProcessor(engine);
                    processor.Rule = columnConfig.Rule;
                    break;
                case "ArithmeticRuleProcessor":
                    processor = ArithmeticRuleProcessor.CreateOrGetProcessor(engine);
                    processor.Rule = columnConfig.Rule;
                    break;
                case "DriverLabelRuleProcessor":
                    processor = DriverLabelRuleProcessor.CreateOrGetProcessor(engine);
                    processor.Rule = columnConfig.Rule;
                    break;
                case "FilesTypeRuleProcessor":
                    processor = FilesTypeRuleProcessor.CreateOrGetProcessor(engine);
                    processor.Rule = columnConfig.Rule;
                    break;
                case "ConditionalRuleProcessor":
                    processor = ConditionalRuleProcessor.CreateOrGetProcessor(engine);
                    processor.Rule = columnConfig.Rule;
                    break;
                case "SimpleReplacementRuleProcess":
                    processor = SimpleReplacementRuleProcess.CreateOrGetProcessor(engine);
                    processor.Rule = columnConfig.Rule;
                    break;
                default:
                    processor = DefaultRuleProcessor.CreateOrGetProcessor(engine);
                    processor.Rule = columnConfig.Rule;
                    break;
            }

            return processor;
        }
    }
}
