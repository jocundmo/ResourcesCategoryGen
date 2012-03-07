using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public interface IFilter
    {
        string Rule { get; set; }
        bool Match(string source);
    }
}
