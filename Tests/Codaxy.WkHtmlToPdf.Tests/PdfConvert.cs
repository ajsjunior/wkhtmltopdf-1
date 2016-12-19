using System;
using System.Diagnostics;
using System.Text;
using System.IO;
using System.Linq;
using System.Threading;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Configuration;

/// <summary>
/// C# wrapper for WKHTMLTOPDF with an easy to use and clean OOP interface.
/// Based on Codaxy.WkHtmlToPdf (https://github.com/codaxy/wkhtmltopdf)
/// Author: Alcir Junior (ajsjunior@gmail.com)
/// </summary>
namespace Codaxy.WkHtmlToPdf
{
    public class PdfConvertException : Exception
    {
        public PdfConvertException(string msg) : base(msg) { }
    }

    public class PdfConvertTimeoutException : PdfConvertException
    {
        public PdfConvertTimeoutException() : base("HTML to PDF conversion process has not finished in the given period") { }
    }

    public class PdfOutput
    {
        public string OutputFilePath { get; set; }
        public Stream OutputStream { get; set; }
        public Action<PdfDocument, byte[]> OutputCallback { get; set; }
    }

    public abstract class PdfObject
    {
        internal abstract void GetCmdArguments(StringBuilder paramsBuilder, WorkEnviroment workEnv);
    }

    public class PdfDocument : PdfObject
    {
        /// <summary>
        /// Collate when printing multiple copies (default TRUE)
        /// </summary>
        public bool? Collate { get; set; }

        /// <summary>
        /// Read and write cookies from and to the supplied cookie jar file
        /// </summary>
        public string CookieJar { get; set; }

        /// <summary>
        /// Number of copies to print into the pdf file (default 1)
        /// </summary>
        public int? Copies { get; set; }

        /// <summary>
        /// Change the dpi explicitly (this has no effect on X11 based systems) (default 96)
        /// </summary>
        public int? Dpi { get; set; }

        /// <summary>
        /// PDF will be generated in grayscale
        /// </summary>
        public bool? Grayscale { get; set; }

        /// <summary>
        /// When embedding images scale them down to this dpi (default 600)
        /// </summary>
        public int? ImageDpi { get; set; }

        /// <summary>
        /// When jpeg compressing images use this quality (default 94)
        /// </summary>
        public int? ImageQuality { get; set; }

        /// <summary>
        /// Generates lower quality pdf/ps. Useful to shrink the result document space
        /// </summary>
        public bool? LowQuality { get; set; }

        /// <summary>
        /// Set the page bottom margin
        /// </summary>
        public double? MarginBottom { get; set; }

        /// <summary>
        /// Set the page left margin (default 10mm)
        /// </summary>
        public double? MarginLeft { get; set; }

        /// <summary>
        /// Set the page right margin (default 10mm)
        /// </summary>
        public double? MarginRight { get; set; }

        /// <summary>
        /// Set the page top margin
        /// </summary>
        public double? MarginTop { get; set; }

        /// <summary>
        /// Set orientation to Landscape or Portrait (default FALSE)
        /// </summary>
        public bool? Landscape { get; set; }

        /// <summary>
        /// Page height
        /// </summary>
        public double? PageHeight { get; set; }

        /// <summary>
        /// Paper sizes (http://doc.qt.io/qt-4.8/qprinter.html#PaperSize-enum)
        /// </summary>
        public enum PaperKind { A0, A1, A2, A3, A4, A5, A6, A7, A8, A9, B0, B1, B2, B3, B4, B5, B6, B7, B8, B9, B10, C5E, Comm10E, DLE, Executive, Folio, Ledger, Legal, Letter, Tabloid, Custom }

        /// <summary>
        /// Set paper size to: A4, Letter, etc. (default A4)
        /// </summary>
        public PaperKind? PageSize { get; set; }

        /// <summary>
        /// Page width
        /// </summary>
        public double? PageWidth { get; set; }

        /// <summary>
        /// Use lossless compression on pdf objects (default TRUE)
        /// </summary>
        public bool? PdfCompression { get; set; }

        /// <summary>
        /// The title of the generated pdf file (The title of the first document is used if not specified)
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Put an outline into the pdf (default TRUE)
        /// </summary>
        public bool? Outline { get; set; }

        /// <summary>
        /// Set the depth of the outline (default 4)
        /// </summary>
        public int? OutlineDepth { get; set; }

        public PageOptions Options = new PageOptions();

        public List<PdfObject> Pages = new List<PdfObject>();

        internal override void GetCmdArguments(StringBuilder paramsBuilder, WorkEnviroment workEnv)
        {
            paramsBuilder.AddArgument(this.Collate, "no-collate", "collate");
            paramsBuilder.AddArgument(this.CookieJar, "cookie-jar");
            paramsBuilder.AddArgument(this.Copies, "copies");
            paramsBuilder.AddArgument(this.Dpi, "dpi");
            paramsBuilder.AddArgument(this.Grayscale, "", "grayscale");
            paramsBuilder.AddArgument(this.ImageDpi, "image-dpi");
            paramsBuilder.AddArgument(this.ImageQuality, "image-quality");
            paramsBuilder.AddArgument(this.LowQuality, "", "lowquality");
            paramsBuilder.AddArgument(this.MarginBottom, "margin-bottom");
            paramsBuilder.AddArgument(this.MarginLeft, "margin-left");
            paramsBuilder.AddArgument(this.MarginRight, "margin-right");
            paramsBuilder.AddArgument(this.MarginTop, "margin-top");
            if (this.Landscape.HasValue)
                paramsBuilder.AddArgument(this.Landscape, "orientation Portrait", "orientation Landscape");
            paramsBuilder.AddArgument(this.PageHeight, "page-height");
            paramsBuilder.AddArgument(this.PageSize, "page-size");
            paramsBuilder.AddArgument(this.PageWidth, "page-width");
            paramsBuilder.AddArgument(this.PdfCompression, "no-pdf-compression", "");
            paramsBuilder.AddArgument(this.Title, "title");
            paramsBuilder.AddArgument(this.Outline, "no-outline", "outline");
            paramsBuilder.AddArgument(this.OutlineDepth, "outline-depth");

            if (Options != null)
                Options.GetCmdArguments(paramsBuilder, workEnv);

            this.Pages.ForEach(obj => obj.GetCmdArguments(paramsBuilder, workEnv));
        }
    }

    public class PageOptions
    {
        /// <summary>
        /// Allow the file or files from the specified folder to be loaded
        /// </summary>
        public List<string> Allow = new List<string>();

        /// <summary>
        /// Do print background (default TRUE)
        /// </summary>
        public bool? Background { get; set; }

        /// <summary>
        /// Bypass proxy for host
        /// </summary>
        public List<string> BypassProxyFor = new List<string>();

        /// <summary>
        /// Web cache directory
        /// </summary>
        public string CacheDir { get; set; }

        /// <summary>
        /// Use this SVG file when rendering checked checkboxes
        /// </summary>
        public string CheckboxCheckedSvg { get; set; }

        /// <summary>
        /// Use this SVG file when rendering unchecked checkboxes
        /// </summary>
        public string CheckboxSvg { get; set; }

        /// <summary>
        /// Set an additional cookie, value should be url encoded.
        /// </summary>
        public Dictionary<string, string> Cookies { get; set; }

        /// <summary>
        /// Set an additional HTTP header
        /// </summary>
        public Dictionary<string, string> CustomHeader { get; set; }

        /// <summary>
        /// Add HTTP headers specified by --custom-header for each resource request.
        /// </summary>
        public bool? CustomHeaderPropagation { get; set; }

        /// <summary>
        /// Show javascript debugging output (default FALSE)
        /// </summary>
        public bool? DebugJavascript { get; set; }

        /// <summary>
        /// Add a default header, with the name of the page to the left, and the page number to the right
        /// </summary>
        public bool? DefaultHeader { get; set; }

        /// <summary>
        /// Set the default text encoding, for input
        /// </summary>
        public Encoding Encoding { get; set; }

        /// <summary>
        /// Make links to remote web pages (default TRUE)
        /// </summary>
        public bool? EnableExternalLinks { get; set; }

        /// <summary>
        /// Turn HTML form fields into pdf form fields (default FALSE)
        /// </summary>
        public bool? EnableForms { get; set; }

        /// <summary>
        /// Do not load or print images (default TRUE)
        /// </summary>
        public bool? Images { get; set; }

        /// <summary>
        /// Make local links (default TRUE)
        /// </summary>
        public bool? EnableInternalLinks { get; set; }

        /// <summary>
        /// Do allow web pages to run javascript (default TRUE)
        /// </summary>
        public bool? EnableJavascript { get; set; }

        /// <summary>
        /// Wait some milliseconds for javascript finish (default 200)
        /// </summary>
        public int? JavascriptDelay { get; set; }

        public enum ErrorHandler { Abort, Ignore, Skip }

        /// <summary>
        /// Specify how to handle pages that fail to load: abort, ignore or skip (default abort)
        /// </summary>
        public ErrorHandler? LoadErrorHandling { get; set; }

        /// <summary>
        /// Specify how to handle media files that fail to load: abort, ignore or skip (default ignore)
        /// </summary>
        public ErrorHandler? LoadMediaErrorHandling { get; set; }

        /// <summary>
        /// Allowed conversion of a local file to read in other local files. (default TRUE)
        /// </summary>
        public bool? EnableLocalFileAccess { get; set; }

        /// <summary>
        /// Minimum font size
        /// </summary>
        public int? MinimumFontSize { get; set; }

        /// <summary>
        /// Include the page in the table of contents and outlines (default TRUE)
        /// </summary>
        public bool? IncludeInOutline { get; set; }

        /// <summary>
        /// Set the starting page number (default 0)
        /// </summary>
        public int? PageOffset { get; set; }

        /// <summary>
        /// HTTP Authentication password
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Enable installed plugins (plugins will likely not work) (default FALSE)
        /// </summary>
        public bool? EnablePlugins { get; set; }

        /// <summary>
        /// Add an additional post field
        /// </summary>
        public Dictionary<string, string> Post { get; set; }

        /// <summary>
        /// Post an additional file
        /// </summary>
        public Dictionary<string, string> PostFile { get; set; }

        /// <summary>
        /// Use print media-type instead of screen (default FALSE)
        /// </summary>
        public bool? PrintMediaType { get; set; }

        /// <summary>
        /// Use a proxy
        /// </summary>
        public string Proxy { get; set; }

        /// <summary>
        /// Use this SVG file when rendering checked radiobuttons
        /// </summary>
        public string RadiobuttonCheckedSvg { get; set; }

        /// <summary>
        /// Use this SVG file when rendering unchecked radiobuttons
        /// </summary>
        public string RadiobuttonSvg { get; set; }

        /// <summary>
        /// Resolve relative external links into absolute links (default TRUE)
        /// </summary>
        public bool? ResolveRelativeLinks { get; set; }

        /// <summary>
        /// Run this additional javascript after the page is done loading
        /// </summary>
        public List<string> RunScript = new List<string>();

        /// <summary>
        /// Enable the intelligent shrinking strategy used by WebKit that makes the pixel/dpi ratio none constant (default TRUE)
        /// </summary>
        public bool? EnableSmartShrinking { get; set; }

        /// <summary>
        /// Stop slow running javascripts (default TRUE)
        /// </summary>
        public bool? StopSlowScripts { get; set; }

        /// <summary>
        /// Link from section header to toc (default FALSE)
        /// </summary>
        public bool? EnableTocBackLinks { get; set; }

        /// <summary>
        /// Specify a user style sheet, to load with every page
        /// </summary>
        public string UserStyleSheet { get; set; }

        /// <summary>
        /// HTTP Authentication username
        /// </summary>
        public string Username { get; set; }

        /// <summary>
        /// Set viewport size if you have custom scrollbars or css attribute overflow to emulate window size
        /// </summary>
        public string ViewportSize { get; set; }

        /// <summary>
        /// Wait until window.status is equal to this string before rendering page
        /// </summary>
        public string WindowStatus { get; set; }

        /// <summary>
        /// Use this zoom factor (default 1)
        /// </summary>
        public double? Zoom { get; set; }

        internal void GetCmdArguments(StringBuilder paramsBuilder, WorkEnviroment workEnv)
        {
            paramsBuilder.AddArgument(this.Allow, "allow");
            paramsBuilder.AddArgument(this.Background, "no-background", "background");
            paramsBuilder.AddArgument(this.BypassProxyFor, "bypass-proxy-for");
            paramsBuilder.AddArgument(this.CacheDir, "cache-dir");
            paramsBuilder.AddArgument(this.CheckboxCheckedSvg, "checkbox-checked-svg");
            paramsBuilder.AddArgument(this.CheckboxSvg, "checkbox-svg");
            paramsBuilder.AddArgument(this.Cookies, "cookie");
            paramsBuilder.AddArgument(this.CustomHeader, "custom-header");
            paramsBuilder.AddArgument(this.CustomHeaderPropagation, "no-custom-header-propagation", "custom-header-propagation");
            paramsBuilder.AddArgument(this.DebugJavascript, "no-debug-javascript", "debug-javascript");
            paramsBuilder.AddArgument(this.DefaultHeader, "", "default-header");
            paramsBuilder.AddArgument(this.Encoding, "encoding");
            paramsBuilder.AddArgument(this.EnableExternalLinks, "disable-external-links", "enable-external-links");
            paramsBuilder.AddArgument(this.EnableForms, "disable-forms", "enable-forms");
            paramsBuilder.AddArgument(this.Images, "no-images", "images");
            paramsBuilder.AddArgument(this.EnableInternalLinks, "disable-internal-links", "enable-internal-links");
            paramsBuilder.AddArgument(this.EnableJavascript, "disable-javascript", "enable-javascript");
            paramsBuilder.AddArgument(this.JavascriptDelay, "javascript-delay");
            paramsBuilder.AddArgument(this.LoadErrorHandling, "load-error-handling");
            paramsBuilder.AddArgument(this.LoadMediaErrorHandling, "load-media-error-handling");
            paramsBuilder.AddArgument(this.EnableLocalFileAccess, "disable-local-file-access", "enable-local-file-access");
            paramsBuilder.AddArgument(this.MinimumFontSize, "minimum-font-size");
            paramsBuilder.AddArgument(this.IncludeInOutline, "exclude-from-outline", "include-in-outline");
            paramsBuilder.AddArgument(this.PageOffset, "page-offset");
            paramsBuilder.AddArgument(this.Password, "password");
            paramsBuilder.AddArgument(this.EnablePlugins, "disable-plugins", "enable-plugins");
            paramsBuilder.AddArgument(this.Post, "post");
            paramsBuilder.AddArgument(this.PostFile, "post-file");
            paramsBuilder.AddArgument(this.PrintMediaType, "no-print-media-type", "print-media-type");
            paramsBuilder.AddArgument(this.Proxy, "proxy");
            paramsBuilder.AddArgument(this.RadiobuttonCheckedSvg, "radiobutton-checked-svg");
            paramsBuilder.AddArgument(this.RadiobuttonSvg, "radiobutton-svg");
            paramsBuilder.AddArgument(this.ResolveRelativeLinks, "keep-relative-links", "resolve-relative-links");
            paramsBuilder.AddArgument(this.RunScript, "run-script");
            paramsBuilder.AddArgument(this.EnableSmartShrinking, "disable-smart-shrinking", "enable-smart-shrinking");
            paramsBuilder.AddArgument(this.StopSlowScripts, "no-stop-slow-scripts", "stop-slow-scripts");
            paramsBuilder.AddArgument(this.EnableTocBackLinks, "disable-toc-back-links", "enable-toc-back-links");
            paramsBuilder.AddUrlArgument(this.UserStyleSheet, workEnv, "user-style-sheet", ".css");
            paramsBuilder.AddArgument(this.Username, "username");
            paramsBuilder.AddArgument(this.ViewportSize, "viewport-size");
            paramsBuilder.AddArgument(this.WindowStatus, "window-status");
            paramsBuilder.AddArgument(this.Zoom, "zoom");
        }
    }

    public class HeaderFooterOptions
    {
        /// <summary>
        /// Centered text
        /// </summary>
        public string Center { get; set; }

        /// <summary>
        /// Set font name (default Arial)
        /// </summary>
        public string FontName { get; set; }

        /// <summary>
        /// Set font size (default 12)
        /// </summary>
        public string FontSize { get; set; }

        /// <summary>
        /// Adds a html text
        /// </summary>
        public string Html { get; set; }

        /// <summary>
        /// Left aligned text
        /// </summary>
        public string Left { get; set; }

        /// <summary>
        /// Display a horizontal line (default FALSE)
        /// </summary>
        public bool? Line { get; set; }

        /// <summary>
        /// Right aligned text
        /// </summary>
        public string Right { get; set; }

        /// <summary>
        /// Spacing between header/footer and content in mm (default 0)
        /// </summary>
        public double? Spacing { get; set; }

        /// <summary>
        /// Replace [name] with value in header and footer
        /// </summary>
        public Dictionary<string, string> Replace { get; set; }

        internal void GetCmdArguments(StringBuilder paramsBuilder, string prefix, WorkEnviroment workEnv)
        {
            paramsBuilder.AddArgument(this.Center, prefix + "-center");
            paramsBuilder.AddArgument(this.FontName, prefix + "-font-name");
            paramsBuilder.AddArgument(this.FontSize, prefix + "-font-size");
            paramsBuilder.AddUrlArgument(this.Html, workEnv, prefix + "-html", ".html");
            paramsBuilder.AddArgument(this.Left, prefix + "-left");
            paramsBuilder.AddArgument(this.Line, "no-" + prefix + "-line", prefix + "-line");
            paramsBuilder.AddArgument(this.Right, prefix + "-right");
            paramsBuilder.AddArgument(this.Spacing, prefix + "-spacing");
            paramsBuilder.AddArgument(this.Replace, "replace");
        }
    }

    public class PdfCover : PdfObject
    {
        public string Html { get; set; }
        public bool ForceHtmlAsContent = false;
        public PageOptions Options = new PageOptions();

        internal override void GetCmdArguments(StringBuilder paramsBuilder, WorkEnviroment workEnv)
        {
            paramsBuilder.Append("cover ");
            paramsBuilder.AddInputArgument(this.Html, workEnv, ".html", ForceHtmlAsContent);

            if (Options != null)
                Options.GetCmdArguments(paramsBuilder, workEnv);
        }
    }

    public class PdfToc : PdfObject
    {
        /// <summary>
        /// Use dotted lines in the toc (default TRUE)
        /// </summary>
        public bool? DottedLines { get; set; }

        /// <summary>
        /// The header text of the toc (default Table of Contents)
        /// </summary>
        public string HeaderText { get; set; }

        /// <summary>
        /// For each level of headings in the toc indent by this length (default 1em)
        /// </summary>
        public string LevelIndentation { get; set; }

        /// <summary>
        /// Link from toc to sections (default TRUE)
        /// </summary>
        public bool? Links { get; set; }

        /// <summary>
        /// For each level of headings in the toc the font is scaled by this factor (default 0.8)
        /// </summary>
        public double? TextSizeShrink { get; set; }

        /// <summary>
        /// Use the supplied xsl style sheet for printing the table of content
        /// </summary>
        public string XslStyleSheet { get; set; }

        public HeaderFooterOptions Header = new HeaderFooterOptions();

        public HeaderFooterOptions Footer = new HeaderFooterOptions();

        internal override void GetCmdArguments(StringBuilder paramsBuilder, WorkEnviroment workEnv)
        {
            paramsBuilder.Append("toc ");
            paramsBuilder.AddArgument(this.DottedLines, "disable-dotted-lines", "");
            paramsBuilder.AddArgument(this.HeaderText, "toc-header-text");
            paramsBuilder.AddArgument(this.LevelIndentation, "toc-level-indentation");
            paramsBuilder.AddArgument(this.Links, "disable-toc-links", "");
            paramsBuilder.AddArgument(this.TextSizeShrink, "toc-text-size-shrink");
            paramsBuilder.AddFileArgument(this.XslStyleSheet, workEnv, "xsl-style-sheet", ".xslt");

            if (Header != null)
                Header.GetCmdArguments(paramsBuilder, "header", workEnv);

            if (Footer != null)
                Footer.GetCmdArguments(paramsBuilder, "footer", workEnv);
        }
    }

    public class PdfPage : PdfObject
    {
        public string Html { get; set; }
        public bool ForceHtmlAsContent = false;
        public PageOptions Options = new PageOptions();
        public HeaderFooterOptions Header = new HeaderFooterOptions();
        public HeaderFooterOptions Footer = new HeaderFooterOptions();

        internal override void GetCmdArguments(StringBuilder paramsBuilder, WorkEnviroment workEnv)
        {
            paramsBuilder.AddInputArgument(this.Html, workEnv, ".html", ForceHtmlAsContent);

            if (Options != null)
                Options.GetCmdArguments(paramsBuilder, workEnv);

            if (Header != null)
                Header.GetCmdArguments(paramsBuilder, "header", workEnv);

            if (Footer != null)
                Footer.GetCmdArguments(paramsBuilder, "footer", workEnv);
        }
    }

    public class PdfConvertEnvironment
    {
        public string TempFolderPath { get; set; }
        public string WkHtmlToPdfPath { get; set; }
        public int? Timeout { get; set; }
    }

    internal class WorkEnviroment
    {
        internal List<string> StdinArguments = new List<string>();
        internal List<string> TempFiles = new List<string>();
        internal string TempFolderPath;

        internal string CreateTempFile(string fileContents, string fileExt)
        {
            var tmpFilePath = Path.Combine(this.TempFolderPath, string.Format("{0}{1}", Guid.NewGuid(), fileExt));
            File.WriteAllText(tmpFilePath, fileContents);
            this.TempFiles.Add(tmpFilePath);
            return tmpFilePath;
        }
    }

    public static class PdfConvert
    {
        private static string GetWkhtmlToPdfExeLocation()
        {
            string filePath;

            filePath = ConfigurationManager.AppSettings["wkhtmltopdf:path"] ?? "";
            filePath = Path.Combine(filePath, @"wkhtmltopdf.exe");
            if (File.Exists(filePath))
                return filePath;

            filePath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles);
            filePath = Path.Combine(filePath, @"wkhtmltopdf\wkhtmltopdf.exe");
            if (File.Exists(filePath))
                return filePath;

            filePath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFilesX86);
            filePath = Path.Combine(filePath, @"wkhtmltopdf\wkhtmltopdf.exe");
            if (File.Exists(filePath))
                return filePath;

            filePath = Path.Combine(@"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe");
            if (File.Exists(filePath))
                return filePath;

            return "wkhtmltopdf.exe";
        }

        /// <summary>
        /// Converts one or more HTML pages into a PDF document.
        /// </summary>
        /// <param name="document">Document definitions.</param>
        /// <param name="output">Output definitions.</param>
        public static void ConvertHtmlToPdf(PdfDocument document, PdfOutput output)
        {
            ConvertHtmlToPdf(document, null, output);
        }

        /// <summary>
        /// Converts one or more HTML pages into a PDF document.
        /// </summary>
        /// <param name="document">Document definitions.</param>
        /// <param name="environment">Environment definitions.</param>
        /// <param name="woutput">Output definitions.</param>
        public static void ConvertHtmlToPdf(PdfDocument document, PdfConvertEnvironment environment, PdfOutput woutput)
        {
            if ((document.Pages == null) || (document.Pages.Count == 0))
                throw new PdfConvertException("You must supply at least one page");

            if (document.Pages.OfType<PdfPage>().Any(p => string.IsNullOrEmpty(p.Html)))
                throw new PdfConvertException("You must supply a HTML string or a URL for all pages");

            if (document.Pages.OfType<PdfCover>().Any(p => string.IsNullOrEmpty(p.Html)))
                throw new PdfConvertException("You must supply a HTML string or a URL for all cover pages");

            if (environment == null)
                environment = new PdfConvertEnvironment();

            if (!environment.Timeout.HasValue)
                environment.Timeout = 60000;

            if (string.IsNullOrEmpty(environment.TempFolderPath))
                environment.TempFolderPath = Path.GetTempPath();

            if (string.IsNullOrEmpty(environment.WkHtmlToPdfPath))
                environment.WkHtmlToPdfPath = GetWkhtmlToPdfExeLocation();

            string outputPdfFilePath;
            bool delete;
            if (woutput.OutputFilePath != null)
            {
                outputPdfFilePath = woutput.OutputFilePath;
                delete = false;
            }
            else
            {
                outputPdfFilePath = Path.Combine(environment.TempFolderPath, string.Format("{0}.pdf", Guid.NewGuid()));
                delete = true;
            }

            // Se for um caminho completo verifica se o executável existe, do contrário considera que ele estará no PATH para rodar corretamente
            if ((environment.WkHtmlToPdfPath.IndexOf(@"\", StringComparison.CurrentCulture) >= 0) && (!File.Exists(environment.WkHtmlToPdfPath)))
                throw new PdfConvertException(string.Format("File '{0}' not found. Check if wkhtmltopdf application is installed.", environment.WkHtmlToPdfPath));

            StringBuilder paramsBuilder = new StringBuilder();
            WorkEnviroment workEnv = new WorkEnviroment() { TempFolderPath = environment.TempFolderPath };
            document.GetCmdArguments(paramsBuilder, workEnv);
            paramsBuilder.AppendFormat("\"{0}\" ", outputPdfFilePath);

            try
            {
                StringBuilder output = new StringBuilder();
                StringBuilder error = new StringBuilder();

                using (Process process = new Process())
                {
                    process.StartInfo.FileName = environment.WkHtmlToPdfPath;
                    process.StartInfo.Arguments = paramsBuilder.ToString().TrimEnd();
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.RedirectStandardOutput = true;
                    process.StartInfo.RedirectStandardError = true;
                    process.StartInfo.RedirectStandardInput = true;
                    process.StartInfo.WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory;

                    using (AutoResetEvent outputWaitHandle = new AutoResetEvent(false))
                    using (AutoResetEvent errorWaitHandle = new AutoResetEvent(false))
                    {
                        DataReceivedEventHandler outputHandler = (sender, e) =>
                        {
                            if (e.Data == null)
                                outputWaitHandle.Set();
                            else
                                output.AppendLine(e.Data);
                        };

                        DataReceivedEventHandler errorHandler = (sender, e) =>
                        {
                            if (e.Data == null)
                                errorWaitHandle.Set();
                            else
                                error.AppendLine(e.Data);
                        };

                        process.OutputDataReceived += outputHandler;
                        process.ErrorDataReceived += errorHandler;

                        Debug.Print("Converting to PDF: {0} {1}", process.StartInfo.FileName, process.StartInfo.Arguments);
                        try
                        {
                            process.Start();

                            process.BeginOutputReadLine();
                            process.BeginErrorReadLine();

                            //TODO:Usar parâmetro --read-args-from-stdin para otimizar geração de vários PDFs em processo batch
                            if (workEnv.StdinArguments.Count > 0) // Se precisa passar algum HTML por stdin
                                using (var stream = process.StandardInput)
                                {
                                    workEnv.StdinArguments.ForEach(input =>
                                    {
                                        byte[] buffer = Encoding.UTF8.GetBytes(input);
                                        stream.BaseStream.Write(buffer, 0, buffer.Length);
                                        stream.WriteLine();
                                    });
                                };

                            if (process.WaitForExit(environment.Timeout.Value) && outputWaitHandle.WaitOne(environment.Timeout.Value) && errorWaitHandle.WaitOne(environment.Timeout.Value))
                            {
                                if (process.ExitCode != 0 && !File.Exists(outputPdfFilePath))
                                    throw new PdfConvertException(string.Format("Html to PDF conversion failed. Wkhtmltopdf output: \r\n{0}\r\nCommand line: {1} {2}", error, process.StartInfo.FileName, process.StartInfo.Arguments));
                            }
                            else
                            {
                                if (!process.HasExited)
                                    process.Kill();

                                throw new PdfConvertTimeoutException();
                            }
                        }
                        finally
                        {
                            process.OutputDataReceived -= outputHandler;
                            process.ErrorDataReceived -= errorHandler;
                        }
                    }
                }

                if (woutput.OutputStream != null)
                {
                    using (Stream fs = new FileStream(outputPdfFilePath, FileMode.Open))
                    {
                        byte[] buffer = new byte[32 * 1024];
                        int read;

                        while ((read = fs.Read(buffer, 0, buffer.Length)) > 0)
                            woutput.OutputStream.Write(buffer, 0, read);
                    }
                }

                if (woutput.OutputCallback != null)
                {
                    byte[] pdfFileBytes = File.ReadAllBytes(outputPdfFilePath);
                    woutput.OutputCallback(document, pdfFileBytes);
                }

            }
            finally
            {
                if (delete && File.Exists(outputPdfFilePath))
                    File.Delete(outputPdfFilePath);

                foreach (var tmpFile in workEnv.TempFiles)
                {
                    if (File.Exists(tmpFile))
                        File.Delete(tmpFile);
                }
            }
        }
    }

    public static class Extensions
    {
        internal static void AddArgument(this StringBuilder builder, object value, string paramName)
        {
            if (value == null)
                return;
            else if (value is Dictionary<string, string>)
                builder.AddRepeatableArgument(value as Dictionary<string, string>, paramName);
            else if (value is List<string>)
                builder.AddRepeatableArgument(value as List<string>, paramName);
            else if (value is Encoding)
                builder.AppendFormat("--{0} \"{1}\" ", paramName, (value as Encoding).WebName);
            //else if (value is Enum)
            //    builder.AppendFormat("--{0} {1} ", paramName, value.ToString().ToLower(CultureInfo.CurrentCulture));
            else if (value is string)
                builder.AppendFormat("--{0} \"{1}\" ", paramName, value);
            else
                builder.AppendFormat("--{0} {1} ", paramName, value);
        }

        internal static void AddArgument(this StringBuilder builder, bool? value, string paramFalse, string paramTrue)
        {
            if (value.HasValue)
                builder.AppendFormat("--{0} ", (value.Value ? paramTrue : paramFalse));
        }

        internal static void AddFileArgument(this StringBuilder builder, string value, WorkEnviroment workEnv, string paramName, string fileExt)
        {
            AddArgument(builder, value, paramName, workEnv, false, fileExt);
        }

        internal static void AddInputArgument(this StringBuilder builder, string value, WorkEnviroment workEnv, string fileExt, bool forceHtmlAsContent)
        {
            AddArgument(builder, value, null, workEnv, forceHtmlAsContent, fileExt);
        }

        private static void AddArgument(this StringBuilder builder, string value, string paramName, WorkEnviroment workEnv, bool forceHtmlAsContent, string fileExt)
        {
            if (!string.IsNullOrEmpty(value))
            {
                if (!string.IsNullOrEmpty(paramName))
                    builder.AppendFormat("--{0} ", paramName);

                if (!forceHtmlAsContent && value.IsUrlOrFilePath())
                    builder.Append(value + " ");
                else
                {
                    //builder.Append("- "); //Extraído de: http://stackoverflow.com/questions/21864382/is-there-any-wkhtmltopdf-option-to-convert-html-text-rather-than-file
                    //workEnv.StdinArguments.Add(value);

                    //if ((fileExt.ToLower(System.Globalization.CultureInfo.CurrentCulture) == ".html") && (!value.StartsWith("<!DOCTYPE html>", true, System.Globalization.CultureInfo.CurrentCulture)))
                    //    value = "<!DOCTYPE html>" + value; // Necessário para funcionar

                    value = workEnv.CreateTempFile(value, fileExt);
                    builder.AppendFormat("\"{0}\" ", value);
                };
            }
        }

        internal static void AddUrlArgument(this StringBuilder builder, string value, WorkEnviroment workEnv, string paramName, string fileExt)
        {
            AddArgument(builder, value, paramName, workEnv, false, fileExt);
        }

        private static void AddRepeatableArgument(this StringBuilder builder, List<string> values, string paramName)
        {
            if ((values != null) && (values.Count > 0))
            {
                //builder.AppendFormat("--{0} ", paramName);
                foreach (var value in values)
                    builder.AppendFormat("--{0} {1} ", paramName, value);
            }
        }

        private static void AddRepeatableArgument(this StringBuilder builder, Dictionary<string, string> values, string paramName)
        {
            if ((values != null) && (values.Count > 0))
            {
                foreach (var value in values)
                    builder.AppendFormat("-{0} {1} {2} ", paramName, value.Key, value.Value);
            }
        }

        public static bool IsUrl(this string str)
        {
            return Regex.IsMatch(str, @"^https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-z]{0,6}\b([-a-zA-Z0-9@:%_\+.~#?&//=]*)$");
        }

        public static bool IsFilePath(this string str)
        {
            return Regex.IsMatch(str, @"^(?:[a-zA-Z]\:|\\\\[\w\.]+\\[\w.$]+)\\(?:[\w]+\\)*\w([\w.])+$");
        }

        public static bool IsUrlOrFilePath(this string str)
        {
            return str.IsUrl() || str.IsFilePath();
        }
    }
}
