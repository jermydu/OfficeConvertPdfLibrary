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

//引入ImageMagick
using ImageMagick;

namespace OfficeConvertPdfLibrary
{
    public class ClassOfficeConvertPdfLibrary
    {
        private const float BaseDpiForPdfConversion = 200f;
        private const int PdfSuperSamplingRatio = 1;
        int pdfPageCount = 0;
        int CurrentOutputFilePathIndex = 0;
        String pngPath = "";

        PowerPoint.Application pptApplication;
        PowerPoint.Presentations presentations;
        PowerPoint.Presentation document;

        //释放ppt
        private void ReleasePowerPoint()
        {
            if (document != null)
            {
                document.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(document);
                document = null;
            }
            if(presentations != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(presentations);
                presentations = null;
            }
            if (presentations != null)
            {
                pptApplication.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(pptApplication);
                presentations = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public int PowerPointConvertPdf(String pptPath,String pdfPath,String pngPath)
        {
            pptApplication = new PowerPoint.Application();
            presentations = pptApplication.Presentations;

            try
            {
                document = presentations.Open(pptPath, ReadOnly: Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse, WithWindow: Office.MsoTriState.msoFalse);
                document.ExportAsFixedFormat(pdfPath, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
            }
            catch
            {
                ReleasePowerPoint();
                return 0;
            }

            ReleasePowerPoint();

            ConvertPdf(pdfPath, pngPath);

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

        //pdf 转 图片
        private void ConvertPdf(String InputFilePath, String OutputFilePath)
        {
            MagickReadSettings settings = new MagickReadSettings();

            float dpi = BaseDpiForPdfConversion;
          
            settings.Density = new Density(dpi * PdfSuperSamplingRatio);
            using (MagickImageCollection images = new MagickImageCollection())
            {
                //添加pdf所有页到collection
                images.Read(InputFilePath, settings);
                pdfPageCount = images.Count;

                //遍历每一页
                foreach (MagickImage image in images)
                {
                    if (PdfSuperSamplingRatio > 1)
                    {
                        image.Scale(new Percentage(100 / PdfSuperSamplingRatio));
                    }
                    pngPath = String.Format("{0}\\png{1}.png", OutputFilePath,CurrentOutputFilePathIndex);
                    ConvertImage(image, true);

                    CurrentOutputFilePathIndex++;
                }
            }
        }

        private void ConvertImage(MagickImage image, bool ignoreScale = false)
        {
            image.Quality = 95;

            image.Write(this.pngPath);
        }
    }
}
