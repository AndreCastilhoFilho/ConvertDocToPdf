using Microsoft.Office.Interop.Word;
using System.Configuration;

namespace convertDocToPdf
{
    public class Program
    {
        public static Microsoft.Office.Interop.Word.Document wordDocument { get; set; }
        static void Main(string[] args)
        {
            var docUrl = ConfigurationManager.AppSettings["docUrl"];
            var pdfUrl = ConfigurationManager.AppSettings["PdfUrl"];

            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            wordDocument = appWord.Documents.Open(@docUrl);
            wordDocument.ExportAsFixedFormat(@pdfUrl, WdExportFormat.wdExportFormatPDF);
        }
    }
}
