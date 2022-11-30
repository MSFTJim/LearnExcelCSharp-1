using System.ComponentModel.Design.Serialization;
using System.Data.Common;
using System.Text.RegularExpressions;
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

            if (ClassifyJava())
                OneAskClassification = "Java";
            else if (ClassifyIntegration())
                OneAskClassification = "Power";
            if (ClassifyCloudNative(Title))
                OneAskClassification = "Power";
            if (ClassifyMisc())
                OneAskClassification = "Misc";
            //else
            //    OneAskClassification = "Classification not set";

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
                    return true;
                }
            }
            

            return false;
        }

        private static bool ClassifyMisc()
        {
            return false;
        }
        private static bool ClassifyCloudNative(string Title)
        {
            int CloudNativeCount = 0; 
            List<string> CN_SingleTerms = new List<string>()
                { "AKS","ARO","ACA", "container","k8s", "kubernetes"};
            //  Container App covered by container

            List<string> CN_AKSTerms = new List<string>()
                { "AKS","kubernetes","Kubernetes","k8s"};
            List<string> CN_ACATerms = new List<string>()
                { "ACA","container app","ContainerApp"};           
               
                foreach (string CNTerm in CN_SingleTerms)
                {
                    if (Title.Contains(CNTerm, StringComparison.CurrentCultureIgnoreCase))
                    {
                        // you made it this far so count it as a CN terms and capture which ever term we found
                        CloudNativeCount++;
                        OneAskClassification = CNTerm;

                        // if count is greater than 1, we call it CN and are done
                        if (CloudNativeCount > 1)
                        {
                            OneAskClassification = "Cloud Native";
                            return false;
                        }
                            

                        // Tagging processing
                        
                        if (Title.Contains("container", StringComparison.CurrentCultureIgnoreCase))
                            if (Title.Contains("container app", StringComparison.CurrentCultureIgnoreCase))
                                OneAskClassification = "ACA";
                            else
                                if (Title.Contains("ContainerApp", StringComparison.CurrentCultureIgnoreCase))
                                OneAskClassification = "ACA";
                            else
                                OneAskClassification = "Cloud Native";

                        if (CN_ACATerms.Any(s => s.Equals(CNTerm, StringComparison.CurrentCultureIgnoreCase)))
                            OneAskClassification = "ACA";

                        if (CN_AKSTerms.Any(s => s.Equals(CNTerm, StringComparison.CurrentCultureIgnoreCase)))
                            OneAskClassification = "AKS";

                        string repattern = @"(cloud).(native)";
                        if (Regex.IsMatch(CNTerm, repattern, RegexOptions.IgnoreCase))
                            OneAskClassification = "Cloud Native";

                    }  // if term contains

                } // foreach terms
             

                //Console.WriteLine("final value: " + OneAskClassification + ". Count: " + CloudNativeCount + ", " + Title);          

            return false;
        } // end cn tagging
        private static bool ClassifyIntegration()
        {
            return false;
        }

        private static bool ClassifyJava()
        {
            return false;
        }

        

        static string OneAskClassification = "classification created";

    } // end class Program

 } // end namespace