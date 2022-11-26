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
            xCel.Application myExcel = new Microsoft.Office.Interop.Excel.Application();
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
            string OneAskClassification = "";
            int CloudNativeCount = 0; // aks, aro, aca, container, cloud native, k8s, kubernetes
            int EventDriven = 0; // Event Hub/Grid, Services bus, event driven
            int IntegrationServerless = 0;  // , APIM, serverless, functions, logic apps 
            int JavaApp = 0;  // java, ASA

            OneAskClassification = ClassifyCloudNative(ref Title, ref OneAskClassification, ref NativeCount);

            OneAskClassification = ClassifyIntegrationServerless(Title);
            OneAskClassification = ClassifyEventDriven(Title);
            OneAskClassification = ClassifyJavaApp(Title);

            

            

            // check cloud event driven
            // Event Hub/Grid, Services bus, event driven
            if (Title.Contains("event", StringComparison.CurrentCultureIgnoreCase))
            {
                oneaskClass = "Event Driven Arch";
                EventDriven++;
            };
            if (Title.Contains("hub", StringComparison.CurrentCultureIgnoreCase))
            {
                oneaskClass = "Event Hub";
                EventDriven++;
            };
            if (Title.Contains("grid", StringComparison.CurrentCultureIgnoreCase))
            {
                oneaskClass = "Event Grid";
                EventDriven++;
            };
            if (Title.Contains("service", StringComparison.CurrentCultureIgnoreCase) && Title.Contains("bus", StringComparison.CurrentCultureIgnoreCase))
            {
                oneaskClass = "Service Bus";
                EventDriven++;
            };
            
            return oneaskClass;

        } // end OneAskClassification

     
        private static string ClassifyEventDriven(string title)
        {
            throw new NotImplementedException();
        }

        private static string ClassifyJavaApp(string title)
        {
            throw new NotImplementedException();
        }

        private static bool DoesThisExist(string WhereToLook, string WhatToLookFor)
        {
            if (WhereToLook.Contains(WhatToLookFor, StringComparison.CurrentCultureIgnoreCase))
                return true;
            else
                return false;
        }


        private static bool DoTheseBothExist(string WhereToLook, string WhatToLookFor1, string WhatToLookFor2)
        {
            return true;
        }

        private static string ClassifyCloudNative(ref string _Title, ref string _OneAskClassification, ref int _CloudNativeCount)
        {
            // check cloud native items
            if (_Title.Contains("AKS", StringComparison.CurrentCultureIgnoreCase))
            {
                _OneAskClassification = "AKS";
                _CloudNativeCount++;
            };

            if (DoesThisExist(_Title,"AKS"))
            {
                _OneAskClassification = "AKS";
                _CloudNativeCount++;
            };

            if (Title.Contains("ARO", StringComparison.CurrentCultureIgnoreCase))
            {
                oneaskClass = "ARO";
                CloudNative++;
            };
            if (Title.Contains("Containers", StringComparison.CurrentCultureIgnoreCase))
            {
                oneaskClass = "Cloud Native";
                CloudNative++;
            };

            if (Title.Contains("Cloud", StringComparison.CurrentCultureIgnoreCase) && Title.Contains("Native", StringComparison.CurrentCultureIgnoreCase))
            {
                oneaskClass = "Cloud Native";
                CloudNative++;
            };

            if (CloudNative > 1)
                oneaskClass = "Cloud Native";



        }
        private static string ClassifyIntegrationServerless(string title)
        {
            throw new NotImplementedException();
        }
    } // end class
    } // end namespace