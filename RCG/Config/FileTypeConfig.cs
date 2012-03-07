using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RCG
{
    public class FileTypeConfig
    {
        public string FileTypeExtension { get; set; }
        public System.IO.SearchOption SearchOption { get; set; }

        public FileTypeConfig()
        {
            // Nothing to do.
        }

        public FileTypeConfig(string fileTypeExtension, System.IO.SearchOption searchOption)
        {
            this.FileTypeExtension = fileTypeExtension;
            this.SearchOption = searchOption;
        }
    }
}
