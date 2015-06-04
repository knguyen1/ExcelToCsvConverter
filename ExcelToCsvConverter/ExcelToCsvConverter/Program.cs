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
            if (args.Length > 0)
                path = args[0];
            else
                path = @"\\Amdc2FAS01A.media.global.loc\AM-US-NYC$\DFS\Shared\Dept-NYC1\ClientServices-NYC1\Compliance Entitlement Report\SONY\MO";
                //path = @"C:\Users\knguyen\Desktop\test2\";

            ExcelToCsvConverter converter = new ExcelToCsvConverter(path);
            converter.ConvertSheet(1);
        }
    }
}
