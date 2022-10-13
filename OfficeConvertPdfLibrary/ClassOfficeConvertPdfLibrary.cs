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

            PowerPoint.Presentations presentations = pptApplication.Presentations;

            PowerPoint.Presentation document = presentations.Open(pptPath, ReadOnly: Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse, WithWindow: Office.MsoTriState.msoFalse);
            if (document == null)
            {
                return 0;
            }
            document.ExportAsFixedFormat(pdfPath, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);

            document.Close();
            pptApplication.Quit();

            //System.Runtime.InteropServices.Marshal.ReleaseComObject(rng);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(document);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(presentations);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(pptApplication);

            //调用GC的垃圾收集方法
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return 1;
        }

        public int WordConvertPdf(String WordPath, String pdfPath)
        {
            Word.Application wordApplication;
            wordApplication = new Word.Application();

            Word.Documents documents = wordApplication.Documents;

            Word.Document document = documents.Open(WordPath, System.Reflection.Missing.Value, true);
            if (document == null)
            {
                return 0;
            }
            document.Activate();
            document.ExportAsFixedFormat(pdfPath, Word.WdExportFormat.wdExportFormatPDF);

            document.Close();
            wordApplication.Quit();

            //System.Runtime.InteropServices.Marshal.ReleaseComObject(rng);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(document);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(documents);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApplication);

            //调用GC的垃圾收集方法
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return 1;
        }

        public int ExcelConvertPdf(String ExcelPath, String pdfPath)
        {
            Excel.Application excelApplication;
            excelApplication = new Excel.Application();

            //wbs 最后释放需要用到
            Excel.Workbooks wbs = excelApplication.Workbooks;
            Excel.Workbook document = wbs.Open(ExcelPath, System.Reflection.Missing.Value, true);
            if (document == null)
            {
                return 0;
            }
            document.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,pdfPath);
            document.Close();
            excelApplication.Quit();


            //释放资源
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(rng);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(document);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wbs);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);

            //调用GC的垃圾收集方法
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return 1;
        }
    }
}
