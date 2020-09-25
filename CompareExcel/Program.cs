using CompareExcelCore;
using System;

namespace CompareExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string superSetName = "SuperSet.xlsx"; //should be user input
            string supperSetSheet = "MV"; //should be user input
            string childSetName = "ChildSet.xlsx"; //should be user input
            string childSetSheet = "MV Config"; //should be user input
            string sErr = "";
            OfficeHelper office = new OfficeHelper();
            sErr = office.CompareFiles(superSetName, supperSetSheet, childSetName, childSetSheet);
            if (sErr == "")
            {
                sErr += "Mission Completed!";
            }
            Console.WriteLine(sErr);
            Console.ReadKey();
        }
    }
}
