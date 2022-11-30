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

            if (!ClassifyFusion(Title))
                if (!ClassifyCloudNative(Title))
                    if (!ClassifyJava(Title))
                        if (!ClassifyIntegration(Title))                    
                            { OneAskClassification = "Classification not set"; };            
            
            //if (ClassifyMisc())
            //    OneAskClassification = "Misc";
            //else
            //    OneAskClassification = "Classification not set";

            return OneAskClassification;

        } // end OneAskClassification

        private static bool ClassifyIntegration(string Title)
        {
            bool ClassifiedAsIntegration = false;
            int IntegrationCount = 0;

            List<string> APIM_Terms = new List<string>()
                { "APIM", "API Management" };
            List<string> SB_Terms = new List<string>()
                { "Service Bus", "ServiceBus"};
            List<string> LA_Terms = new List<string>()
                { "Logic Apps","LogicApps"};

            List<string> Integration_Terms = new List<string>()
                { "event", "Functions"};
            Integration_Terms.AddRange(APIM_Terms);
            Integration_Terms.AddRange(LA_Terms);
            Integration_Terms.AddRange(SB_Terms);

            foreach (string IntegrationTerm in Integration_Terms)
            {
                if (Title.Contains(IntegrationTerm, StringComparison.CurrentCultureIgnoreCase))
                {
                    ClassifiedAsIntegration = true;
                    OneAskClassification = IntegrationTerm;
                    IntegrationCount++;
                    if (IntegrationCount > 1)
                    {
                        OneAskClassification = "Integration";
                        break;
                    }

                    if (APIM_Terms.Any(s => s.Equals(IntegrationTerm, StringComparison.CurrentCultureIgnoreCase)))
                        OneAskClassification = "APIM";
                    else
                    if (SB_Terms.Any(s => s.Equals(IntegrationTerm, StringComparison.CurrentCultureIgnoreCase)))
                        OneAskClassification = "Service Bus";
                    else
                    if (LA_Terms.Any(s => s.Equals(IntegrationTerm, StringComparison.CurrentCultureIgnoreCase)))
                        OneAskClassification = "Logic Apps";
                    else
                    if (Title.Contains("Event", StringComparison.CurrentCultureIgnoreCase) && Title.Contains("Hub", StringComparison.CurrentCultureIgnoreCase))
                        OneAskClassification = "Event Hubs";
                    else
                    if (Title.Contains("Event", StringComparison.CurrentCultureIgnoreCase) && Title.Contains("Grid", StringComparison.CurrentCultureIgnoreCase))
                        OneAskClassification = "Event Grid";
                    else
                    if (Title.Contains("Event", StringComparison.CurrentCultureIgnoreCase) && Title.Contains("Driven", StringComparison.CurrentCultureIgnoreCase))
                        OneAskClassification = "Event Driven Arch";
                    

                } // end if contains

            } // end foreach



            return ClassifiedAsIntegration;
        }

        private static bool ClassifyJava(string Title)
        {
            bool ClassifiedAsJava = false;

            if (Title.Contains("spring", StringComparison.CurrentCultureIgnoreCase))
            { OneAskClassification = "Spring";
              ClassifiedAsJava = true;
            }
            else
                if (Title.Contains("java", StringComparison.CurrentCultureIgnoreCase))
            {
                OneAskClassification = "Java";
                ClassifiedAsJava = true;
            }          


            return ClassifiedAsJava;

        }


        private static bool ClassifyFusion(string Title)
        {

            bool ClassifiedAsFusion = false;
            
            List<string> Fusion_Terms = new List<string>()
                { "Power", "Fusion", "RPA", "LC/NC", "Low Code"  };

            foreach (string Fusion_Term in Fusion_Terms)
            {
                if (Title.Contains(Fusion_Term, StringComparison.CurrentCultureIgnoreCase))
                {
                    OneAskClassification = "LC/NC";
                    ClassifiedAsFusion = true;
                    break;
                }
            }
            

            return ClassifiedAsFusion;
        }

        private static bool ClassifyMisc()
        {
            return false;
        }
        private static bool ClassifyCloudNative(string Title)
        {
            bool ClassifiedAsCN = false;

            int CloudNativeCount = 0;
            int AROCount = 0;


            List<string> CN_AKSTerms = new List<string>()
                { "AKS","kubernetes","Kubernetes","k8s"};
            List<string> CN_AROTerms = new List<string>()
                { "ARO","redhat","red hat","openshift","open shift"};
            List<string> CN_ACATerms = new List<string>()
                { "ACA","container app","ContainerApp"};

            List<string> CN_Terms = new List<string>()
                { "container"};

            CN_Terms.AddRange(CN_AKSTerms);
            CN_Terms.AddRange(CN_AROTerms);
            CN_Terms.AddRange(CN_ACATerms);


            //  Container App covered by container

            foreach (string CNTerm in CN_Terms)
                {
                    if (Title.Contains(CNTerm, StringComparison.CurrentCultureIgnoreCase))
                    {
                        // you made it this far so count it as a CN terms and capture which ever term we found
                        ClassifiedAsCN = true;
                        CloudNativeCount++;
                        OneAskClassification = CNTerm;

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


                    if (CN_AROTerms.Any(s => s.Equals(CNTerm, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        AROCount++;
                        if (AROCount > 1)
                            CloudNativeCount--;
                        OneAskClassification = "ARO";
                    }
                    string repattern = @"(cloud).(native)";
                    if (Regex.IsMatch(CNTerm, repattern, RegexOptions.IgnoreCase))
                        OneAskClassification = "Cloud Native";

                    // if count is greater than 1, we call it CN and are done
                    if (CloudNativeCount > 1)
                    {
                        OneAskClassification = "Cloud Native";
                        break;
                    }

                }  // if term contains

                } // foreach terms             

                //Console.WriteLine("final value: " + OneAskClassification + ". Count: " + CloudNativeCount + ", " + Title);          

                return ClassifiedAsCN;

        } // end cn tagging
      

      

        

        static string OneAskClassification = "classification created";

    } // end class Program

 } // end namespace