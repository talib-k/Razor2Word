using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Razor2Word.Utils
{
    public abstract class BaseViewForWord : WebViewPage
    {
        public WordHelpers Word { get; set; }

        public override void InitHelpers()
        {
            base.InitHelpers();
            Word = new WordHelpers<object>(base.ViewContext, this);
        }

    }

    public abstract class BaseViewForWord<TModel> : BaseViewForWord
    {
        public new WordHelpers<TModel> Word { get; set; }

        public override void InitHelpers()
        {
            base.InitHelpers();
            Word = new WordHelpers<TModel>(base.ViewContext, this);
        }
    }

}