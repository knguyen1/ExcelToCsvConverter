using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;

namespace ExcelToCsvConverter
{
    static class ExtensionMethods
    {
        public static IEnumerable<FileInfo> GetFilesByExtensions(this DirectoryInfo dir, params string[] extensions)
        {
            if (extensions == null)
                throw new ArgumentNullException("Extensions cannot be null!");

            IEnumerable<FileInfo> files = dir.EnumerateFiles();
            return files.Where(f => extensions.Contains(f.Extension));
        }
    }
}
