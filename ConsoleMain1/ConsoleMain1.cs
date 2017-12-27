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
        private static double[,] arvot = new double[300, 4];

        public static void Main()
        {
            LueExcelista();
            LaskeKeskiarvot(arvot);
            Console.ReadKey();
        }

        public static double[,] LueExcelista()
        {

            //Tässä alustetaan Excel-tiedosto käyttöön
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbooks xlWorkbookS = xlApp.Workbooks;
            Excel.Workbook xlWorkbook = xlWorkbookS.Open(@"D:\repos\ConsoleMain1\ConsoleMain1\Resources\Kuntatutka.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            
            int arvotrow = 0;
            int arvotcol = 0;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            

            //Tässä laitetaan Excel tarkastamaan solut rivistä 2 alkaen, sarakkeet 47-50 eli AU-AX ja tallennetaan ne taulukkoon
            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 47; j <= 50; j++)
                {
                    arvotrow = i - 2;
                    arvotcol = j - 47;
                    double arvo = Convert.ToDouble(xlRange.Cells[i, j].Value2);
                    arvot[arvotrow, arvotcol] = arvo;
                }
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //Objektien vapauttaminen, jotta excel-prosessit saadaan suljettua taustalta
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //Sulkeminen ja vapauttaminen
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlWorkbookS);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return arvot;

        }

        public static void LaskeKeskiarvot(double[,] arvot)
        {

            double temp = 0;

            for (int j = 0; j <= 0; j++)
            {
                for (int i = 0; i < arvot.GetLength(0); i++)
                {
                    temp = temp+arvot[i, j];
                }

                double keskiarvo = temp / 295;
                Console.WriteLine("AU-sarakkeen arvojen summa: " + temp.ToString());
                Console.WriteLine("AU-sarakkeen arvojen keskiarvo: " + keskiarvo.ToString());
            }

            temp = 0;

            for (int j = 1; j <= 1; j++)
            {
                for (int i = 0; i < arvot.GetLength(0); i++)
                {
                    temp = temp + arvot[i, j];
                }

                double keskiarvo = temp / 295;
                Console.WriteLine("AV-sarakkeen arvojen summa: " + temp.ToString());
                Console.WriteLine("AV-sarakkeen arvojen keskiarvo: " + keskiarvo.ToString());
            }

            temp = 0;

            for (int j = 2; j <= 2; j++)
            {
                for (int i = 0; i < arvot.GetLength(0); i++)
                {
                    temp = temp + arvot[i, j];
                }

                double keskiarvo = temp / 295;
                Console.WriteLine("AW-sarakkeen arvojen summa: " + temp.ToString());
                Console.WriteLine("AW-sarakkeen arvojen keskiarvo: " + keskiarvo.ToString());
            }

            temp = 0;

            for (int j = 3; j <= 3; j++)
            {
                for (int i = 0; i < arvot.GetLength(0); i++)
                {
                    temp = temp + arvot[i, j];
                }

                double keskiarvo = temp / 295;
                Console.WriteLine("AX-sarakkeen arvojen summa: " + temp.ToString());
                Console.WriteLine("AX-sarakkeen arvojen keskiarvo: " + keskiarvo.ToString());
            }

        }
    }
}