﻿using System;
using System.Diagnostics;
using System.Text;
using System.IO;
using System.Web;

namespace Codaxy.WkHtmlToPdf
{
    public class PdfConvertException : Exception
    {
        public PdfConvertException(String msg) : base(msg) { }
    }

    public class PdfConvertTimeoutException : PdfConvertException
    {
        public PdfConvertTimeoutException() : base("HTML to PDF conversion process has not finished in the given period.") { }
    }

    public class PdfOutput
    {
        public String OutputFilePath { get; set; }
        public Stream OutputStream { get; set; }
        public Action<PdfDocument, byte[]> OutputCallback { get; set; }
    }

    public class PdfDocument
    {
        public String Url { get; set; }
        public String Html { get; set; }
        public String HeaderUrl { get; set; }
        public String FooterUrl { get; set; }
        public String HeaderLeft { get; set; }
        public String HeaderCenter { get; set; }
        public String HeaderRight { get; set; }
        public String FooterLeft { get; set; }
        public String FooterCenter { get; set; }
        public String FooterRight { get; set; }
        public object State { get; set; }
    }

    public class PdfConvertEnvironment
    {
        public String TempFolderPath { get; set; }
        public String WkHtmlToPdfPath { get; set; }
        public int Timeout { get; set; }
        public bool Debug { get; set; }
    }

    public class PdfConvert
    {
        static PdfConvertEnvironment _e;

        public static PdfConvertEnvironment Environment
        {
            get
            {
                if (_e == null)
                    _e = new PdfConvertEnvironment
                    {
                        TempFolderPath = Path.GetTempPath(),
                        WkHtmlToPdfPath = GetWkhtmlToPdfExeLocation(),
                        Timeout = 60000
                    };
                return _e;
            }
        }

        private static string GetWkhtmlToPdfExeLocation()
        {
            string programFilesPath = System.Environment.GetEnvironmentVariable("ProgramFiles");
            string filePath = Path.Combine(programFilesPath, @"wkhtmltopdf\wkhtmltopdf.exe");

            if (File.Exists(filePath))
                return filePath;

            string programFilesx86Path = System.Environment.GetEnvironmentVariable("ProgramFiles(x86)");
            filePath = Path.Combine(programFilesx86Path, @"wkhtmltopdf\wkhtmltopdf.exe");

            if (File.Exists(filePath))
                return filePath;

            filePath = Path.Combine(programFilesPath, @"wkhtmltopdf\bin\wkhtmltopdf.exe");
            if (File.Exists(filePath))
                return filePath;

            return Path.Combine(programFilesx86Path, @"wkhtmltopdf\bin\wkhtmltopdf.exe");
        }

        public static void ConvertHtmlToPdf(PdfDocument document, PdfOutput output)
        {
            ConvertHtmlToPdf(document, null, output);
        }

        public static void ConvertHtmlToPdf(PdfDocument document, PdfConvertEnvironment environment, PdfOutput woutput)
        {
            if (document.Url == "-" && document.Html == null)
                throw new PdfConvertException(
                    String.Format("You must supply a HTML string, if you have enterd the url: {0}", document.Url)
                );
            if (environment == null)
                environment = Environment;

            String outputPdfFilePath;
            bool delete;
            if (woutput.OutputFilePath != null)
            {
                outputPdfFilePath = woutput.OutputFilePath;
                delete = false;
            }
            else
            {
                outputPdfFilePath = Path.Combine(environment.TempFolderPath, String.Format("{0}.pdf", Guid.NewGuid()));
                delete = true;
            }

            if (!File.Exists(environment.WkHtmlToPdfPath))
                throw new PdfConvertException(String.Format("File '{0}' not found. Check if wkhtmltopdf application is installed.", environment.WkHtmlToPdfPath));

            ProcessStartInfo si;

            StringBuilder paramsBuilder = new StringBuilder();
            paramsBuilder.Append("--page-size A4 ");
            //paramsBuilder.Append("--redirect-delay 0 "); not available in latest version
            if (!string.IsNullOrEmpty(document.HeaderUrl))
            {
                paramsBuilder.AppendFormat("--header-html {0} ", document.HeaderUrl);
                paramsBuilder.Append("--margin-top 25 ");
                paramsBuilder.Append("--header-spacing 5 ");
            }
            if (!string.IsNullOrEmpty(document.FooterUrl))
            {
                paramsBuilder.AppendFormat("--footer-html {0} ", document.FooterUrl);
                paramsBuilder.Append("--margin-bottom 25 ");
                paramsBuilder.Append("--footer-spacing 5 ");
            }

            if (!string.IsNullOrEmpty(document.HeaderLeft))
                paramsBuilder.AppendFormat("--header-left \"{0}\" ", document.HeaderLeft);

            if (!string.IsNullOrEmpty(document.FooterCenter))
                paramsBuilder.AppendFormat("--header-center \"{0}\" ", document.HeaderCenter);

            if (!string.IsNullOrEmpty(document.FooterCenter))
                paramsBuilder.AppendFormat("--header-right \"{0}\" ", document.HeaderRight);

            if (!string.IsNullOrEmpty(document.FooterLeft))
                paramsBuilder.AppendFormat("--footer-left \"{0}\" ", document.FooterLeft);

            if (!string.IsNullOrEmpty(document.FooterCenter))
                paramsBuilder.AppendFormat("--footer-center \"{0}\" ", document.FooterCenter);

            if (!string.IsNullOrEmpty(document.FooterCenter))
                paramsBuilder.AppendFormat("--footer-right \"{0}\" ", document.FooterRight);

            paramsBuilder.AppendFormat("\"{0}\" \"{1}\"", document.Url, outputPdfFilePath);

            si = new ProcessStartInfo();
            si.CreateNoWindow = !environment.Debug;
            si.FileName = environment.WkHtmlToPdfPath;
            si.Arguments = paramsBuilder.ToString();
            si.UseShellExecute = false;
            si.RedirectStandardError = !environment.Debug;
            si.RedirectStandardInput = true;

            try
            {
                using (var process = new Process())
                {
                    process.StartInfo = si;
                    process.Start();
                    if (document.Html != null)
                        using (var stream = process.StandardInput)
                        {
                            byte[] buffer = Encoding.UTF8.GetBytes(document.Html);
                            stream.BaseStream.Write(buffer, 0, buffer.Length);
                            stream.WriteLine();
                        }
                    if (!process.WaitForExit(environment.Timeout))
                        throw new PdfConvertTimeoutException();

                    if (!File.Exists(outputPdfFilePath))
                    {
                        if (process.ExitCode != 0)
                        {
                            var error = si.RedirectStandardError ? process.StandardError.ReadToEnd() : String.Format("Process exited with code {0}.", process.ExitCode);
                            throw new PdfConvertException(String.Format("Html to PDF conversion of '{0}' failed. Wkhtmltopdf output: \r\n{1}", document.Url, error));
                        }

                        throw new PdfConvertException(String.Format("Html to PDF conversion of '{0}' failed. Reason: Output file '{1}' not found.", document.Url, outputPdfFilePath));
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
                        woutput.OutputCallback(document, File.ReadAllBytes(outputPdfFilePath));
                    }
                }
            }
            finally
            {
                if (delete && File.Exists(outputPdfFilePath))
                    File.Delete(outputPdfFilePath);
            }
        }
    }
}