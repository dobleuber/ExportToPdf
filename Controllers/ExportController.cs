using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web.Http;
using System.Web.Hosting;
using Nustache.Core;
using System.Dynamic;
using Codaxy.WkHtmlToPdf;
using System.Net.Http.Headers;
using Carvajal.Cosmos.Domain.DTO;
using System.IO;
using System.Web;
using System.Web.Http.Cors;

namespace Carvajal.Cosmos.WebAPI.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class ExportController : ApiController
    {
        public string GetHtmlString(ExportDescriptor descriptor) 
        {
            var simpleObject = descriptor.ExportObject;
            if (string.IsNullOrWhiteSpace(simpleObject))
            {
                return string.Empty;
            }

            dynamic dynJson = JsonConvert.DeserializeObject(simpleObject);

            var sbResult = new StringBuilder();

            foreach (var field in descriptor.FieldList)
            {

                AddLine(sbResult);

                var dataValue = GetValue(dynJson, field.Property, field.Format);

                if (field.DetailDescriptor == null || field.DetailDescriptor.FieldList.Count == 0)
                {
                    var exportedField = new
                    {
                        Label = field.Label,
                        Value = dataValue
                    };

                    sbResult.Append(Render.FileToString(GetPath("Templates/SingleField.html"), exportedField));
                }
                else
                {                    
                    var detailDescription = field.DetailDescriptor;
                    var columns = new List<dynamic>();
                    var rows = new List<dynamic>();

                    var sbRow = new StringBuilder();
                    foreach (var detailRow in dataValue.Children())
                    {
                        AddLine(sbRow);
                        var sbCell = new StringBuilder();
                        foreach (var column in detailDescription.FieldList)
                        {
                            AddLine(sbCell);
                            sbCell.Append(Render.FileToString(GetPath("Templates/DetailCell.html"), new { Detail = GetValue(detailRow, column.Property, column.Format)}));
                        }

                        sbRow.AppendFormat(Render.FileToString(GetPath("Templates/DetailRow.html"), new { Detail = sbCell.ToString() }));
                    }

                    var exportedDetailData = new { TableLabel = field.Label, Columns = detailDescription.FieldList, TBody = sbRow.ToString() };

                    sbResult.Append(Render.FileToString(GetPath("Templates/DetailsField.html"), exportedDetailData));
                }
            }

            return Render.FileToString(GetPath("Templates/ExportLayout.html"), new { body = sbResult.ToString() });
        }

        [HttpPost, Route("api/ExportToHtml")]
        public HttpResponseMessage ExportToHtml(ExportDescriptor descriptor)
        {
            var htmlToExport = GetHtmlString(descriptor);
            HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
            response.Content = new StringContent(htmlToExport);
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
            return response;
        }

        [HttpPost, Route("api/ExportToPdf")]
        public string ExportToPdf(ExportDescriptor descriptor)
        {
            var exportPath = Path.Combine(HttpRuntime.AppDomainAppPath, "Exports");

            if (!Directory.Exists(exportPath))
            {
                Directory.CreateDirectory(exportPath);
            }

            var outputPdf =  new PdfOutput
            {
                OutputFilePath = Path.Combine(exportPath, descriptor.FileName + ".pdf")
            };

            var htmlToExport = GetHtmlString(descriptor);

            PdfConvert.ConvertHtmlToPdf(new PdfDocument { Url = "-", 
                Html = htmlToExport, 
                HeaderUrl = GetPath("Templates/Header.html"), 
                FooterUrl = GetPath("Templates/Footer.html") }, outputPdf);

            return string.Format("exports/{0}", Path.GetFileName(outputPdf.OutputFilePath));
        }

        protected dynamic GetValue(dynamic dynJson, string property, string format = null)
        {
            var containerObject = dynJson;
            if (property.Contains('.'))
            {
                var props = property.Split('.');
                int i;
                for (i = 0; i < props.Length - 1; i++)
                {
                    containerObject = containerObject[props[i]];
                }

                property = props[i];
            }

            return format == null ? containerObject[property] : containerObject[property].ToString(format);
        }

        private static string GetPath(string path)
        {
            return File.Exists(path) ? path : Path.Combine(HttpRuntime.AppDomainAppPath, path);
        }

        private static void AddLine(StringBuilder sbIn)
        {
            if (sbIn.Length > 0)
            {
                sbIn.AppendLine();
            }
        }
    }
}
