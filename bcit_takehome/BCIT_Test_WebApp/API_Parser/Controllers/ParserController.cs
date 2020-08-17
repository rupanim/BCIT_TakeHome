using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Http;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Web;
using System;
using System.Net;
using System.Net.Http;

namespace API_Parser.Controllers
{
    public class ParserController : ApiController
    {
        private Dictionary<string, string> GetMarkers()
        {
            return (
                new Dictionary<string, string>()
                {
                    { "#heading1", "<h1>" },
                    { "/heading1", "</h1>" },
                    { "#heading2", "<h2>" },
                    { "/heading2", "</h2>" },
                    { "#heading3", "<h3>" },
                    { "/heading3", "</h3>" },
                    { "#outcome", "<p><div class=\"panel panel-info panel-body\">" },
                    { "/outcome", "</div></p>" },
                    { "#image", "<br /><br /><img src>" },
                    { "/image", "</img><br /><br />" },
                    { "#imageTitle", "<br /><label><b>" },
                    { "/imageTitle", "</b></label>" },
                    { "#subheadingaccordion", "<button data-toggle=\"collapse\" data-target=\"#accordion_id_div\">" },
                    { "/subheadingaccordion", "</button>" },
                    { "#accordionInfo", "<div id=\"accordion_id_section\" class=\"collapse\"><p>" },
                    { "/accordionInfo", "</div></p>" },
                    {"#reflection", "<div class=\"panel panel-default panel-body\"><div class=\"panel-body\"><b>" },
                    {"/reflection", "</b></div></div>" },
                    {"#key-point", "<div class=\"panel panel-info panel-body heading3\"><ul>" },
                    {"/key-point", "</ul></div>" },
                     { "#item", "<li>" },
                    { "/item", "</li>" },
                    { "#reading", "<div class=\"panel panel-default\">" },
                    { "/reading", "</div>" },
                    {"#subheading", "<h4 class=\"subheading\">" },
                    {"/subheading", "</h4>" },
                    {"#accordion", "<br /><div class=\"accordion\">" },
                    {"/accordion", "</div>" }
                }
            );
        }

        private Dictionary<string, string> GetMarkersPreview()
        {
            return (
                new Dictionary<string, string>()
                {
                    { "#heading1", "<h1>" },
                    { "/heading1", "</h1>" },
                    { "#heading2", "<h2>" },
                    { "/heading2", "</h2>" },
                    { "#heading3", "<h3>" },
                    { "/heading3", "</h3>" },
                    { "#outcome", "<br />" },
                    { "/outcome", "<br />" },
                    { "#image", "<br /><br /><b>Missing Image!!</b>" },
                    { "/image", "" },
                    { "#imageTitle", "<br /><label><b>" },
                    { "/imageTitle", "</b></label><br />" },
                    { "#subheadingaccordion", "<br />" },
                    { "/subheadingaccordion", "<br />" },
                    { "#accordionInfo", "<br />" },
                    { "/accordionInfo", "<br />" },
                    {"#reflection", "<b>" },
                    {"/reflection", "</b>" },
                    {"#key-point", "<div class=\"heading3\"><ul>" },
                    {"/key-point", "</ul></div>" },
                     { "#item", "<li>" },
                    { "/item", "</li>" },
                    { "#reading", "<br />" },
                    { "/reading", "" },
                    {"#subheading", "<h4 class=\"subheading\">" },
                    {"/subheading", "</h4>" },
                    {"#accordion", "<br /><div class=\"accordion\">" },
                    {"/accordion", "</div>" }
                }
            );
        }

        private bool FileValidation(string filePath)
        {
            return File.Exists(filePath);
        }

        private string GetHtmlHeader()
        {
            return @"<html>
                    <head>
                    <meta charset = ""utf-8"" ><meta name = ""viewport"" content = ""width=device-width, initial-scale=1"" ><link rel = ""stylesheet"" href = ""https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css"" >
                    <script src = ""https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"" >
                    </script ><script src = ""https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"">
                    </script >
                    <style>
                    body
                    {font:18px""OpenSans"",HelveticaNeue,Helvetica,Arial,sans-serif;line-height:1.5;text-rendering:optimizeLegibility;}
                    .heading2,.subheading{padding-left:1em;}
                    .heading3{background-color:#e6eeff;}
                    .accordion{background-color:#eee;color:#444;cursor:pointer;padding:18px;width:100%;border:none;text-align:left;outline:none;font-size:15px;transition:0.4s;}

                    </style>

                    </head><body><div class=""container"">";
        }

        private string GetHtmlFooter()
        {
            return "</body></html>";
        }

        [HttpGet]
        public HttpResponseMessage Get(string file)
        {
            string responseStr = ParseDocxToHtml(file, out HttpStatusCode responseStatusCode);
            
            return this.Request.CreateResponse(responseStatusCode, responseStr);
        }      

            private string ParseDocxToHtml(string filepath, out HttpStatusCode responseStatusCode)
        { 
            StringBuilder sb = new StringBuilder();
            StringBuilder sb_preview = new StringBuilder();

            try
            {
                if (!FileValidation(filepath))
                {
                    responseStatusCode = HttpStatusCode.NotFound;
                    return "File not found.";
                }

                Dictionary<string, string> markers = GetMarkers();
                Dictionary<string, string> previewMarkers = GetMarkersPreview();
                

                string fileFolderName = Path.GetDirectoryName(filepath);
                string fileName = Path.GetFileName(filepath).Replace(".docx", "");
                string imagesFolderName = fileFolderName + "\\" + fileName + "_ParsedImages";
                if (Directory.Exists(imagesFolderName))
                {
                    Directory.Delete(imagesFolderName, true);
                }                

                Directory.CreateDirectory(imagesFolderName);

                using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, false))
                {
                    Queue<string> imagesList = new Queue<string>();
                    int accordionCount = 0;

                    var imageParts = wordprocessingDocument.MainDocumentPart.ImageParts;
                    foreach (var imagePart in imageParts)
                    {
                        var uri = imagePart.Uri;
                        var filename = uri.ToString().Split('/').Last();
                        var stream = wordprocessingDocument.Package.GetPart(uri).GetStream();
                        Bitmap b = new Bitmap(stream);
                        string imageLoc = imagesFolderName + "\\" + filename;
                        b.Save(imageLoc);
                        imagesList.Enqueue(imageLoc);
                    }
                    var paragraphs = wordprocessingDocument.MainDocumentPart.RootElement.Descendants<Paragraph>();
                    foreach (var paragraph in paragraphs)
                    {
                        if (previewMarkers.TryGetValue(paragraph.InnerText.Trim(), out string p))
                        {
                            sb_preview.AppendLine(p);
                        }
                        else {
                            sb_preview.AppendLine(paragraph.InnerText.Trim());
                        }
                        
                        if (markers.TryGetValue(paragraph.InnerText.Trim(), out string el))
                        { 
                            if (el.Contains("accordion_id_div"))
                            {
                                accordionCount++;
                                el = el.Replace("accordion_id_div", "accordion_" + accordionCount);
                            }

                            if (el.Contains("accordion_id_section"))
                            {
                                el = el.Replace("accordion_id_section", "accordion_" + accordionCount);
                            }

                            if (el.Contains("<img src>") && imagesList.Count != 0)
                            {                                
                                sb.AppendLine(el.Replace("src", "src = \"" + imagesList.Dequeue() + "\""));
                            }
                            else
                            {                                
                                sb.AppendLine(el);
                            }
                        }
                        else
                        {                            
                            sb.AppendLine(paragraph.InnerText.Trim());
                        }
                    }
                }       
                File.WriteAllText(fileFolderName + "\\" + fileName + ".htm", GetHtmlHeader() + sb.ToString() + GetHtmlFooter());
                responseStatusCode = HttpStatusCode.OK;
            }
            catch (Exception ex)
            {
                sb.Clear();
                sb_preview.Clear();
                sb_preview.AppendLine("Parsing Error: Unable to parse file. Details: ");
                sb_preview.AppendLine(ex.ToString());
                responseStatusCode = HttpStatusCode.BadRequest;
            }         

             return (sb_preview.ToString());
            
        }

    }
}
