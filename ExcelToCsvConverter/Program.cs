using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToCsvConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            string path;

            if (args == null)
                throw new ArgumentNullException("Path must be defined.");

            if (args.Length > 0)
                path = args[0];
            else
                throw new ArgumentNullException("Path must be defined.");

            ExcelToCsvConverter converter = new ExcelToCsvConverter(path);
            converter.ConvertSheet(1); //converts the sheet at position 1
        }
    }
}
