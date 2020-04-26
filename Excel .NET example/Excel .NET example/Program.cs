using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Excel.NET_example
{
    class Program
    {
        static Dictionary<string, string> Notas = new Dictionary<string, string>();

        static void Main(string[] args)
        {
            var fi = new FileInfo(@"xls/Notas.xlsx");
            ExcelPackage package = new ExcelPackage(fi);
            ExcelWorkbook xlWorkbook = package.Workbook;

            ExcelWorksheet WSNotas = xlWorkbook.Worksheets[1];
            CargarNotas(WSNotas);

            ExcelWorksheet WSBase = xlWorkbook.Worksheets[2];
            PlaySong(WSBase);

            Console.ReadKey();
        }

        private static void CargarNotas(ExcelWorksheet WSNotas)
        {
            if (WSNotas.Dimension == null) return;
            if (WSNotas.Dimension.End == null) return;

            int iRowCnt = WSNotas.Dimension.End.Row;
            if (iRowCnt == 0) return;

            int i = 0;
            try
            {
                for (i = 2; i <= iRowCnt; i++)
                {
                    if (WSNotas.Cells[i, 5].Value == null || WSNotas.Cells[i, 5].Value.ToString().Trim().Length == 0)
                        continue;
                    string idx = WSNotas.Cells[i, 6].Value.ToString();
                    //Double dVal = Double.Parse(WSNotas.Cells[i, 5].Value.ToString());
                    //int val = Convert.ToInt32(dVal);
                    string val = WSNotas.Cells[i, 5].Value.ToString();
                    Notas[idx] = val;
                }
            }
            catch (Exception ex)
            { Console.WriteLine("Error: " + ex.Message + " (" + ex.GetType().Name + ")"); }
        }

        private static void PlaySong(ExcelWorksheet WSBase)
        {
            if (WSBase.Dimension == null) return;
            if (WSBase.Dimension.End == null) return;

            int iRowCnt = 20;//WSBase.Dimension.End.Row;
            if (iRowCnt == 0) return;

            int dTono = int.Parse(WSBase.Cells[1, 1].Value.ToString());

            int i = 0; // para el bucle

            try
            {
                for (i = 2; i <= iRowCnt; i++)
                {
                    if (WSBase.Cells[i, 5].Value == null || WSBase.Cells[i, 5].Value.ToString().Trim().Length == 0)
                        break;

                    string nota = WSBase.Cells[i, 5].Value.ToString();
                    int delay = int.Parse(WSBase.Cells[i, 3].Value.ToString());
                    int duracion = int.Parse(WSBase.Cells[i, 4].Value.ToString());

                    int tDly = Convert.ToInt32((delay * dTono) / 4);
                    int tDur = Convert.ToInt32((tDly * duracion) / 10);

                    Console.Beep(5000, 1000);

                    // de Arduino
                    //tone(11, iTono, tDur);
                    //delay(tDly);
                    //noTone(11);

                    Console.WriteLine(i.ToString() + "::" + nota);
                    Thread.Sleep(tDly);
                }
            }
            catch (Exception ex)
            { Console.WriteLine("Error: " + ex.Message + " (" + ex.GetType().Name + ")"); }
        }
    }
}
