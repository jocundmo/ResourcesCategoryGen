using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class NamingMatchFilter : IFilter
    {
        #region IFilter Members

        public string Rule { get; set; }

        public bool Match(string source)
        {
            return source == Rule;
        }

        #endregion
    }
}
