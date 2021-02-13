using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WordToPdfConverter.Models;

namespace WordToPdfConverter.Controllers
{
    public class HomeController : Controller
    {
        public async System.Threading.Tasks.Task<ActionResult> Index(string button)
        {
            if (button == null)
                return View();
            ViewBag.Message = string.Empty;
            if (Request.Form.Files != null)
            {
                
                string extension = Path.GetExtension(Request.Form.Files[0].FileName).ToLower();
               
                if (extension == ".doc" || extension == ".docx" || extension == ".docm"
                   || extension == ".xml" )
                {
                    MemoryStream stream = new MemoryStream();                   
                    Request.Form.Files[0].CopyTo(stream);
                    try
                    {
                        
                        WordDocument document = new WordDocument(stream, FormatType.Automatic);
                        stream.Dispose();
                        stream = null;

                        DocIORenderer render = new DocIORenderer();
                        Syncfusion.Pdf.PdfDocument pdf = render.ConvertToPDF(document);

                        MemoryStream memoryStream = new MemoryStream();
                        pdf.Save(memoryStream);
                        render.Dispose();
                        document.Close();
                        pdf.Close();
                        memoryStream.Position = 0;

                        return File(memoryStream, "application/pdf", "WordToPDF.pdf");
                    }
                    catch (Exception ex)
                    {
                        ViewBag.Message = string.Format("The input document could not be processed completely.");
                    }
                }
                else
                {
                    ViewBag.Message = string.Format("Choose a Word format document to convert to PDF");
                }
            }
            
            return View();

        }
    }
}
