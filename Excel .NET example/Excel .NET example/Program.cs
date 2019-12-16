using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.NET_example
{
    class Program
    {
        static void Main(string[] args)
        {
            var fi = new FileInfo(@"Notas.xlsx");
            ExcelPackage package = new ExcelPackage(fi);
            ExcelWorkbook xlWorkbook = package.Workbook;
            ExcelWorksheet WSBase = xlWorkbook.Worksheets[1];

            int iRowCnt = WSBase.Dimension.End.Row;
            int iCol = 2;// WSBase.Dimension.End.Column;

            int i = 0;
            try
            {
                for (i = 2; i <= 4; i++)
                {
                    if (WSBase.Cells[i, iCol].Value == null || WSBase.Cells[i, iCol].Value.ToString().Trim().Length == 0)
                        continue;

                    Console.WriteLine("dataval: " + WSBase.Cells[i, iCol].Value);
                    /*
                    DataRow drw = dt.NewRow();
                    // ---------------- 0 - NUMERO DOCUMENTO IDENTIDAD ----------------
                    string celda = (WSBase.Cells[i, 1].Value ?? string.Empty).ToString();
                    if (celda.Trim().Length == 0)
                        regListasError("1", "Documento de Identidad no Ingresado", ref idxCol, ref errorFila);
                    else drw[0] = celda.Trim().ToUpper();
                    // ---------------- 1 - CODIGO DE ENVIO ----------------
                    drw[1] = CodigoEnvio;

                    if (errorFila.Count > 0)
                    {
                        string[] eFila = { i.ToString(), string.Join(", ", errorFila.ToArray()) };
                        if (idxCol.Count > 0)
                            eFila = eFila.Concat(new string[] { string.Join(":", idxCol.ToArray()) }).ToArray();
                        listaError.Add(eFila);
                        continue;
                    }
                    dt.Rows.Add(drw);*/
                }
            }
            catch (Exception ex)
            { Console.WriteLine("Error: " + ex.Message + " (" + ex.GetType().Name + ")"); }
            Console.ReadKey();
        }
    }
}
