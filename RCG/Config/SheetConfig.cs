using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class SheetConfig
    {
        private List<ColumnConfig> _columns = new List<ColumnConfig>();
        private List<LocationConfig> _locations = new List<LocationConfig>();
        private List<FilterConfig> _filters = new List<FilterConfig>();
        private List<FormatterConfig> _formatters = new List<FormatterConfig>();

        public string Name { get; set; }
        public string Mode { get; set; }
        public int MaxRowCount { get; set; }
        public bool Enabled { get; set; }
        public bool RefMode { get; set; }
        public List<ColumnConfig> Columns { get { return _columns; } }
        public List<LocationConfig> Locations { get { return _locations; } }
        public List<FilterConfig> Filters { get { return _filters; } }
        public List<FormatterConfig> Formatters { get { return _formatters; } }

        public SheetConfig(string name, bool enabled, bool refMode, string mode, int maxRowCount)
        {
            this.Name = name;
            this.Enabled = enabled;
            this.RefMode = refMode;
            this.Mode = mode;
            this.MaxRowCount = maxRowCount;
        }

        public SheetConfig(string name, string mode, int maxRowCount)
            : this(name, true, true, mode, maxRowCount)
        {
            // Nothing to do.
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
