using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;


namespace GenerateScript
{
    class Program
    {
        const string path = @"C:\Users\Jenell\Downloads\ranges.xlsx";
        const string per0txt = @"C:\Users\Jenell\Desktop\per0.sim";
        const string per25txt = @"C:\Users\Jenell\Desktop\per25.sim";
        const string per50txt = @"C:\Users\Jenell\Desktop\per50.sim";
        const string per75txt = @"C:\Users\Jenell\Desktop\per75.sim";
        const string per100txt = @"C:\Users\Jenell\Desktop\per100.sim";

        static string[] writeTo = new string[] { per0txt, per25txt, per50txt, per75txt, per100txt };


        static void Main(string[] args)
        {
            System.IO.FileInfo fi = new System.IO.FileInfo(path);

            const int maxRow = 107;

            using (ExcelPackage xlPackage = new ExcelPackage(fi))
            {
                // get the first worksheet in the workbook
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[1];

                string CPA = worksheet.Cells[3, 1].Value.ToString();

                int ind = 6;

                for(int i = 0;i<=4;i++)
                {
                    //Set up the file
                    using (StreamWriter fs = new StreamWriter(writeTo[i])) {

                        fs.WriteLine("EAL:");

                        for (int row = 3; row <= 107; row++)
                        {
                            string cpa = worksheet.Cells[row, 1].Value.ToString();
                            string val = worksheet.Cells[row, ind].Value.ToString();
                            string writeMe = "EAI:" + cpa + "/" + val + "//1";

                            fs.WriteLine(writeMe);
                        }

                        fs.WriteLine("USM:");

                    }

                    ind += 2;

                }



            } // the using statement calls Dispose() which closes the package.
            
          
        }
    }
}
