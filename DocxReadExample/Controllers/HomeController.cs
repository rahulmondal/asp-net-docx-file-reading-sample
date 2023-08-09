using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxReadExample.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            var filepath = Server.MapPath("~/App_Data/test.docx");
            using (var wordDocument = WordprocessingDocument.Open(filepath, false))
            {
                ViewBag.DocBody = wordDocument.MainDocumentPart.Document.Body.InnerText;
            }
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";
            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";
            return View();
        }
    }
}