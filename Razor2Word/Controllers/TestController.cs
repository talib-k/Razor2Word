using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Windows;
using Razor2Word.Utils;

namespace Razor2Word.Controllers
{
    public class TestController : BaseWordGenerationController
    {
        public TestController()
        {
            PageConfiguration = PageDescription.A4.Landscape.WithMargins(new PageMargins(10));
        }

        public ActionResult Index()
        {
            return Word();
        }

    }
}
