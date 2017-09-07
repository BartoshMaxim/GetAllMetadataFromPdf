using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetAllMetadataFromPdf.Core
{
    public class PdfManager : IPdfManager<string>
    {
        public string CreateExcelWithPdfMetadata(string pdfPath)
        {
            if (File.Exists(pdfPath))
            {
                IDictionary<string, string> metadata = GetAllPdfMetadata(pdfPath);
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                string directoryFullPath = Path.GetDirectoryName(pdfPath);
                string xlsFileName = Path.GetFileNameWithoutExtension(pdfPath) + ".xlsx";
                string xlsFullPath = directoryFullPath + "\\" + xlsFileName;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);               

                int position = 1;
                foreach (var item in metadata)
                {
                    byte[] keyValueBytes = Encoding.Default.GetBytes(item.Key);
                    xlWorkSheet.Cells[position, 1] = Encoding.UTF8.GetString(keyValueBytes);

                    byte[] valueBytes = Encoding.Default.GetBytes(item.Value);
                    xlWorkSheet.Cells[position, 2] = Encoding.UTF8.GetString(valueBytes);
                    position++;
                }
                xlWorkBook.SaveAs(xlsFullPath);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                return xlsFullPath;
            }
            throw new FileNotFoundException();
        }

        public IDictionary<string, string> GetAllPdfMetadata(string pdfPath)
        {
            if (File.Exists(pdfPath))
            {
                PdfReader pdfReader = new PdfReader(pdfPath);
                return pdfReader.Info;
            }
            throw new FileNotFoundException();
        }
    }
}
