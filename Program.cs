using System.ComponentModel.Design.Serialization;
using System.Data.Common;
using xCel = Microsoft.Office.Interop.Excel;

namespace ConsoleExcel2
{
    internal class Program
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

                    myWorkssheet.Cells[row, col] = OneAskClassification(myWorkssheet.Cells[row, col + 1].Value2);

                else

                    myWorkssheet.Cells[row, col] = "Null Title";

            }

            myWorkbook.Save();
            myWorkbook.Close();

            Console.WriteLine("Row count: "+lastRow);
            Console.WriteLine("End!");


        } // end main
 

       
        private static string OneAskClassification(string Title)
        {
            string oneaskClass = "classification not set";            
            
            return oneaskClass;

        } // end OneAskClassification
    
    } // end class Program

 } // end namespace