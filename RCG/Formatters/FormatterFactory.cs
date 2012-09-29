using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class FormatterFactory
    {
        public static IFormatter GetFormatter(string formatterType, string rule, string formatString, GenProcessor engine)
        {
            IFormatter formatter = null;

            switch (formatterType)
            {
                case "NumberCompareFormatter":
                    formatter = NumberCompareFormatter.CreateNewProcessor(engine) as IFormatter;
                    formatter.Rule = rule;
                    formatter.FormatString = formatString;
                    break;
                case "DateTimeCompareFormatter":
                    formatter = DateTimeCompareFormatter.CreateNewProcessor(engine) as IFormatter;
                    formatter.Rule = rule;
                    formatter.FormatString = formatString;
                    break;
                case "AppendedItemFormatter":
                    formatter = AppendedItemFormatter.CreateNewProcessor(engine) as IFormatter;
                    formatter.Rule = rule;
                    formatter.FormatString = formatString;
                    break;
                case "UpdatedItemFormatter":
                    formatter = UpdatedItemFormatter.CreateNewProcessor(engine) as IFormatter;
                    formatter.Rule = rule;
                    formatter.FormatString = formatString;
                    break;
                case "RefreshedItemFormatter":
                    formatter = RefreshedItemFormatter.CreateNewProcessor(engine) as IFormatter;
                    formatter.Rule = rule;
                    formatter.FormatString = formatString;
                    break;
                case "DuplicatedItemFormatter":
                    formatter = DuplicatedItemFormatter.CreateNewProcessor(engine) as IFormatter;
                    formatter.Rule = rule;
                    formatter.FormatString = formatString;
                    break;
                case "DeletedItemFormatter":
                    formatter = DeletedItemFormatter.CreateNewProcessor(engine) as IFormatter;
                    formatter.Rule = rule;
                    formatter.FormatString = formatString;
                    break;
                default:
                    throw new Exception(string.Format("Formatter {0} is not recognized", formatterType));
            }

            return formatter;
        }
    }
}
