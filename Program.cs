using System;
using GNAlibrary;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Configuration;
using System.IO;
using System.Collections.Generic;

namespace MatchPrismToTOR
{
    internal class Program
    {
        static void Main(string[] args)
        {
            gnaToolbox gna = new gnaToolbox();


            string strExcelPath = System.Configuration.ConfigurationManager.AppSettings["ExcelPath"];
            string strExcelFile = System.Configuration.ConfigurationManager.AppSettings["ExcelFile"];
            string strPrismWorksheet = System.Configuration.ConfigurationManager.AppSettings["PrismWorksheet"];
            string strTORworksheet = System.Configuration.ConfigurationManager.AppSettings["TORworksheet"];
            string strMatchedWorksheet = System.Configuration.ConfigurationManager.AppSettings["MatchedWorksheet"];
            double dblMaximumSeparation = Convert.ToDouble(System.Configuration.ConfigurationManager.AppSettings["MaximumSeparation"]);
            string strWorkingFile = strExcelPath + strExcelFile;

            string strPrismName = "x";
            double dblPrismN;
            double dblPrismE;
            double dblPrismH;
            string strToRName = "x";
            double dblToRN;
            double dblToRE;
            double dblToRH;

            string strFormula1 = "";
            string strFormula2 = "";



            //================[Declare variables]=============================================================

            // Console settings
            Console.OutputEncoding = System.Text.Encoding.Unicode;

           //==============[



            //================[Main program]===========================================================================

            // Set the EPPlus license
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            // Welcome message
            gna.WelcomeMessage("Match prisms to TOR survey");


            // Set useful global variables
            double dblS = 0.0;


            // Read in the monitoring prisms

            // open the existing workbook
            var fiExcelWorkbook = new FileInfo(strWorkingFile);

            // Read in the monitoring prism data from the prism worksheet

            Int16 iRow = 1;

            //Prism prism = new  Prism();
 

            using (var package = new ExcelPackage(fiExcelWorkbook))
            {

                // Read in the monitoring prisms
                Console.WriteLine("1. Extract the monitoring prisms");
                var worksheet = package.Workbook.Worksheets[strPrismWorksheet];
                var prism = new List<Prism>();

                do
                {
                    
                    strPrismName = Convert.ToString(worksheet.Cells[iRow, 1].Value);
                    dblPrismE = Convert.ToDouble(worksheet.Cells[iRow, 2].Value);
                    dblPrismN = Convert.ToDouble(worksheet.Cells[iRow, 3].Value);
                    dblPrismH = Convert.ToDouble(worksheet.Cells[iRow, 4].Value);

                    prism.Add(new Prism() { Name = strPrismName, E = dblPrismE, N = dblPrismN, H = dblPrismH });

                    Console.WriteLine(strPrismName);


                    iRow++;
                    strPrismName = Convert.ToString(worksheet.Cells[iRow, 1].Value);

                } while (strPrismName != "");

                // Read in the ToR observations
                Console.WriteLine("2. Extract the Top of Rail readings");
                var tor = new List<ToR>();
               iRow = 1;
               worksheet = package.Workbook.Worksheets[strTORworksheet];

                do
                {

                    strToRName = Convert.ToString(worksheet.Cells[iRow, 1].Value);
                    dblToRE = Convert.ToDouble(worksheet.Cells[iRow, 2].Value);
                    dblToRN = Convert.ToDouble(worksheet.Cells[iRow, 3].Value);
                    dblToRH = Convert.ToDouble(worksheet.Cells[iRow, 4].Value);

                    tor.Add(new ToR() { Name = strToRName, E = dblToRE, N = dblToRN, H = dblToRH });

                    Console.WriteLine(strToRName);
                    
                    iRow++;
                    strToRName = Convert.ToString(worksheet.Cells[iRow, 1].Value);

                } while (strToRName != "");

               // add the "used" marker which when sorted ends up at the bottom of the list
               tor.Add(new ToR() { Name = "zzzzzzz", E = 0.0, N = 0.0, H = 0.0 });


                // Step through the prism data finding the closest ToR reading
                Console.WriteLine("3. Match ToR to Prisms");

                int iNumber =tor.Count;
                var combined = new List<Combined>();

                for (int iPrismNumber = 0; iPrismNumber < prism.Count; iPrismNumber++)
                {

                    // sort the ToR data
                    tor.Sort(delegate (ToR x, ToR y)
                    {
                        return x.Name.CompareTo(y.Name);
                    });

                    strPrismName= prism[iPrismNumber].Name;
                    dblPrismN = prism[iPrismNumber].N;
                    dblPrismE = prism[iPrismNumber].E;
                    dblPrismH = prism[iPrismNumber].H;

                    double dblMinDistance = 9999999.00;      // records the minimum distance encountered during the filtering
                    int iToRindex=0;                        // records the associated ToR index
                    dblS = 0.0;

                    int j = 0;
                    strToRName = tor[0].Name;
                    while (strToRName != "zzzzzzz")
                    {
                        dblToRN = tor[j].N;
                        dblToRE = tor[j].E;
                        dblS = Math.Pow((dblPrismN - dblToRN),2) + Math.Pow((dblPrismE - dblToRE),2);
                        dblS = Math.Pow(dblS, 0.5);
                        if (dblS < dblMinDistance)
                        {
                            iToRindex=j;
                            dblMinDistance=dblS;
                        }
                        j++;
                        strToRName = tor[j].Name;
                    }

                    strToRName = tor[iToRindex].Name;
                    dblToRN = tor[iToRindex].N;
                    dblToRE = tor[iToRindex].E;
                    dblToRH = tor[iToRindex].H;

                    //Console.WriteLine(strPrismName + " : " + dblMinDistance);
                    //Console.WriteLine(strPrismName + " " + dblPrismE + " " + dblPrismN + " " + dblPrismH);
                    //Console.WriteLine(strToRName + " " + dblToRE + " " + dblToRN + " " + dblToRH);
                    //Console.ReadLine();



                    if (dblMinDistance < dblMaximumSeparation)
                    {
                        tor[iToRindex].Name = "zzzzzzz";

                        combined.Add(new Combined()
                        {
                            PrismName = strPrismName,
                            PrismE = dblPrismE,
                            PrismN = dblPrismN,
                            PrismH = dblPrismH,
                            TORName = strToRName,
                            TORE = dblToRE,
                            TORN = dblToRN,
                            TORH = dblToRH
                        });

                    }
                    else
                    {

                        combined.Add(new Combined()
                        {
                            PrismName = strPrismName,
                            PrismE = dblPrismE,
                            PrismN = dblPrismN,
                            PrismH = dblPrismH,
                            TORName = "Missing ToR Reading",
                            TORE = 0.0,
                            TORN = 0.0,
                            TORH = 0.0
                        });
                    }
                }


                combined.Add(new Combined()
                {
                    PrismName = "TheEnd",
                    PrismE = 0.0,
                    PrismN = 0.0,
                    PrismH = 0.0,
                    TORName = "TheEnd",
                    TORE = 0.0,
                    TORN = 0.0,
                    TORH = 0.0
                });



                // write to the combined worksheet
                Console.WriteLine("4. Write to spreadsheet");
                worksheet = package.Workbook.Worksheets[strMatchedWorksheet];
                iRow = 1;
                int iIndex = iRow-1;
                strPrismName = combined[0].PrismName;
                while (strPrismName != "TheEnd")
                {
                    strFormula1 = "=((F" + Convert.ToString(iRow) + "-B" + Convert.ToString(iRow) + ")^2 + (G" + Convert.ToString(iRow) + "-C" + Convert.ToString(iRow) + ")^2)^0.5";
                    strFormula2 = "=(H" + Convert.ToString(iRow) + ")-(D" + Convert.ToString(iRow) + ")";

                    worksheet.Cells[iRow, 1].Value = combined[iIndex].PrismName;
                    worksheet.Cells[iRow, 2].Value = combined[iIndex].PrismE;
                    worksheet.Cells[iRow, 3].Value = combined[iIndex].PrismN;
                    worksheet.Cells[iRow, 4].Value = combined[iIndex].PrismH;
                    worksheet.Cells[iRow, 5].Value = combined[iIndex].TORName;


                    if (Convert.ToString(combined[iIndex].TORName) != "Missing ToR Reading")
                    {
                        worksheet.Cells[iRow, 9].Formula = strFormula1;
                        worksheet.Cells[iRow, 10].Formula = strFormula2;
                        worksheet.Cells[iRow, 6].Value = combined[iIndex].TORE;
                        worksheet.Cells[iRow, 7].Value = combined[iIndex].TORN;
                        worksheet.Cells[iRow, 8].Value = combined[iIndex].TORH;
                    }


                    iRow++;
                    iIndex = iRow - 1;
                    strPrismName = combined[iIndex].PrismName;
                }

                worksheet.Cells["A1:J3000"].Style.Numberformat.Format = "0.000";

                worksheet.Calculate();
                package.Save();
            }

            Console.WriteLine("5. Done");
            Console.WriteLine("");
            Console.WriteLine("Press enter...");
            Console.ReadKey();
        }
    }
}
