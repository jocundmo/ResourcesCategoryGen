using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.ObjectModel;

namespace RCG
{
    public class LocationConfig
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public bool Enabled { get; set; }
        public bool IncludeFolder { get; set; }
        public Collection<FileTypeConfig> IncludeFileTypes { get; private set; }

        public LocationConfig(string name, string path, bool enabled, bool includeFolder)
        {
            this.Name = name;
            this.Path = path;
            this.Enabled = enabled;
            this.IncludeFolder = includeFolder;
            this.IncludeFileTypes = new Collection<FileTypeConfig>();
        }

        public LocationConfig()
            : this (string.Empty, string.Empty, true, true)
        {
            // Nothing to do.
        }

        public LocationConfig(string path)
            : this(path, path, true, true)
        {
            // Nothing to do.
        }

        public override string ToString()
        {
            return Path;
        }
    }
}
