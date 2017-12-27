using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Yle
{
    public class Excel_Testi
    {
        private static Object[,] arvot = new Object[300, 90];
        
        //Tässä alustetaan Excel-tiedosto käyttöön
        private static Excel.Application xlApp = new Excel.Application();
        private static Excel.Workbooks xlWorkbookS = xlApp.Workbooks;
        private static Excel.Workbook xlWorkbook = xlWorkbookS.Open(@"D:\repos\ConsoleMain1\ConsoleMain1\Resources\Kuntatutka.xlsx");
        private static Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
        private static Excel.Range xlRange = xlWorksheet.UsedRange;
        private static int rowCount = xlRange.Rows.Count;
        private static int colCount = xlRange.Columns.Count;

        public static void Main()
        {
            LaskeKeskiarvot();
            Console.Write("\r\nSyötä arvioitava maakunta (muista isot alkukirjaimet):");
            string maakunta = Console.ReadLine();
            LaskeMaakunnat(maakunta);

            Console.WriteLine("\r\nPaina mitä vaan poistuaksesi.");
            Console.ReadKey();

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //Objektien vapauttaminen, jotta excel-prosessit saadaan suljettua taustalta
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //Työkirjan sulkeminen ja vapauttaminen
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlWorkbookS);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

        }

        //Tälle funktiolle syötetään lähtöarvona laskettava taulukko, josta se laskee sarakkeiden keskiarvot
        public static void LaskeKeskiarvot()
        {

            double temp = 0;
            int maara = 0;

            for (int j = 47; j <=50; j++)
            {
                maara = 0;
                temp = 0;

                for (int i = 2; i <= rowCount; i++)
                {
                    temp = temp+Convert.ToDouble(xlRange.Cells[i, j].Value2);
                    maara++;
                }

                double keskiarvo = temp / 295;
                keskiarvo = Math.Round(keskiarvo, 2);
                
                Console.WriteLine("\r\n" + (j).ToString() + ". sarakkeen arvojen keskiarvo koko maassa: " + keskiarvo.ToString());
            }
            Console.WriteLine("\r\nKuntia koko Suomessa: " + maara.ToString());
        }
        
        //Tämä funktio ottaa syötteenä halutun maakunnan nimen, ja antaa sitä koskevien tunnuslukujen keskiarvot
        public static void LaskeMaakunnat(string s)
        {
            double temp = 0;
            int maara = 0;

            for (int j = 47; j <= 50; j++)
            {
                maara = 0;
                temp = 0;

                for (int i = 2; i <= rowCount; i++)
                {
                    if (xlRange.Cells[i, 4].Value2 == s && xlRange.Cells[i, j].Value2 != null)
                    {
                        temp = temp + Convert.ToDouble(xlRange.Cells[i, j].Value2);
                        maara++;
                    }
                }

                double keskiarvo = temp / maara;
                keskiarvo = Math.Round(keskiarvo, 2);

                Console.WriteLine("\r\n" + (j).ToString() + ". sarakkeen arvojen keskiarvo" + "(" + s + "): " + keskiarvo.ToString());
            }
            Console.WriteLine("\r \nKuntia alueella " + s + ": " + maara.ToString());
        }
    }
}