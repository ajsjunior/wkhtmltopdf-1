using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Codaxy.WkHtmlToPdf.Tests
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.InputEncoding = Encoding.UTF8;

            PdfConvert.ConvertHtmlToPdf(new PdfDocument
            {
                Pages = { new PdfPage { Html = "http://www.codaxy.com" } }
            }, new PdfOutput
            {
                OutputFilePath = "codaxy.pdf"
            });

            PdfConvert.ConvertHtmlToPdf(new PdfDocument
            {
                Pages = { new PdfPage{
                    Html = "http://www.codaxy.com",
                    Header = { Left = "[title]", Right = "[date] [time]" },
                    Footer = { Center = "Page [page] of [topage]" } } }
            }, new PdfOutput
            {
                OutputFilePath = "codaxy_hf.pdf"
            });

            PdfConvert.ConvertHtmlToPdf(new PdfDocument
            {
                Pages = { new PdfPage { Html = "<html><h1>test</h1></html>" } }
            }, new PdfOutput
            {
                OutputFilePath = "inline.pdf"
            });

            PdfConvert.ConvertHtmlToPdf(new PdfDocument
            {
                Pages = { new PdfPage { Html = "<html><h1>測試</h1></html>" } }
            }, new PdfOutput
            {
                OutputFilePath = "inline_cht.pdf"
            });

            //PdfConvert.ConvertHtmlToPdf("http://tweakers.net", "tweakers.pdf");
        }
    }
}
