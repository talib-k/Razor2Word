using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NotesFor.HtmlToOpenXml;
using Razor2Word.Utils;

namespace Razor2Word.Controllers
{
    public abstract class BaseWordGenerationController : Controller
    {
        public PageDescription PageConfiguration { get; set; }

        protected ActionResult Word(string viewName = null, object model = null)
        {
            var fileContent = GenerateWord(viewName, model);
            return File(fileContent,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        }

        public byte[] GenerateWord(string viewName, object model)
        {
            var wordGenerator = new WordGenerator(FindView(viewName))
            {
                PageConfiguration = PageConfiguration
            };
            return wordGenerator.GenerateWord(model);
        }

        private IView FindView(string viewName)
        {
            if (viewName == null)
            {
                viewName = Convert.ToString(ControllerContext.RouteData.Values["action"]);
            }
            return ViewEngines.Engines.FindView(ControllerContext, viewName, null).View;
        }
    }
}
