using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//引入office
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficeConvertPdfLibrary
{
    public class ClassOfficeConvertPdfLibrary
    {
        public int PowerPointConvertPdf(String pptPath,String pdfPath)
        {
            PowerPoint.Application pptApplication;
            pptApplication = new PowerPoint.Application();

            PowerPoint.Presentation document = pptApplication.Presentations.Open(pptPath, ReadOnly: Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse, WithWindow: Office.MsoTriState.msoFalse);
            if (document == null)
            {
                return 0;
            }
            document.ExportAsFixedFormat(pdfPath, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);

            document.Close();
            document = null;
            pptApplication.Quit();
            pptApplication = null;

            return 1;
        }

        public int WordConvertPdf(String WordPath, String pdfPath)
        {
            Word.Application wordApplication;
            wordApplication = new Word.Application();

            Word.Document document = wordApplication.Documents.Open(WordPath, System.Reflection.Missing.Value, true);
            if (document == null)
            {
                return 0;
            }
            document.Activate();
            document.ExportAsFixedFormat(pdfPath, Word.WdExportFormat.wdExportFormatPDF);

            document.Close();
            document = null;
            wordApplication.Quit();
            wordApplication = null;

            return 1;
        }

        public int ExcelConvertPdf(String ExcelPath, String pdfPath)
        {
            Excel.Application excelApplication;
            excelApplication = new Excel.Application();

            Excel.Workbook document = excelApplication.Workbooks.Open(ExcelPath, System.Reflection.Missing.Value, true);
            if (document == null)
            {
                return 0;
            }
            document.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,pdfPath);

            document.Close();
            document = null;
            excelApplication.Quit();
            excelApplication = null;

            return 1;
        }
    }
}
