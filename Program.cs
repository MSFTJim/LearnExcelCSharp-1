using System.ComponentModel.Design.Serialization;
using System.Data.Common;
using xCel = Microsoft.Office.Interop.Excel;

namespace ConsoleExcel2
{
     class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start!");
            string OneAskFile = "C:\\Users\\jamesvac\\Documents\\OneAskData3.xlsx";
            //xCel.Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            xCel.Application myExcel = new();
            xCel.Workbook myWorkbook;
            xCel.Worksheet myWorkssheet;
            
            

            myWorkbook = myExcel.Workbooks.Open(OneAskFile);
            myWorkssheet = myWorkbook.Worksheets[1];
            int lastRow = myWorkssheet.Cells.SpecialCells(xCel.XlCellType.xlCellTypeLastCell).Row;
           
            //int row = 2;
            int col = 4;

            for (int row = 2; row <= lastRow; row++)
            {
                if (myWorkssheet.Cells[row, col + 1].Value2 != null)

                    myWorkssheet.Cells[row, col] = ClassifyOneAsk(myWorkssheet.Cells[row, col + 1].Value2);

                else

                    myWorkssheet.Cells[row, col] = "Null Title";

            }

            myWorkbook.Save();
            myWorkbook.Close();

            Console.WriteLine("Row count: "+lastRow);
            Console.WriteLine("End!");

        } // end main
       
        private static string ClassifyOneAsk(string Title)
        {
            OneAskClassification = "Classification started";            

            if (ClassifyFusion(Title))
                OneAskClassification = "Power";
                
            //if (ClassifyJava())
            //        OneAskClassification = "Java";
            //        else if (ClassifyIntegration())
            //            OneAskClassification = "Power";
            //            if (ClassifyCloudNative())
            //                OneAskClassification = "Power";
            //                if (ClassifyMisc())
            //                   OneAskClassification = "Classification not set";
            //                   else
            //                        OneAskClassification = "Classification not set";

            return OneAskClassification;

        } // end OneAskClassification


        private static bool ClassifyFusion(string Title)
        {

            List<string> Fusion_Terms = new List<string>()
                { "Power", "Fusion", "RPA", "LC/NC", "Low Code"  };

            foreach (string Fusion_Term in Fusion_Terms)
            {
                if (Title.Contains(Fusion_Term, StringComparison.CurrentCultureIgnoreCase))
                {
                    OneAskClassification = "LC/NC";
                    break;
                }
            }
            

            return false;
        }

        private static bool ClassifyMisc()
        {
            return true;
        }
        private static bool ClassifyCloudNative()
        {
            return true;
        }
        private static bool ClassifyIntegration()
        {
            return true;
        }

        private static bool ClassifyJava()
        {
            return true;
        }

        

        static string OneAskClassification = "classification created";

    } // end class Program

 } // end namespace