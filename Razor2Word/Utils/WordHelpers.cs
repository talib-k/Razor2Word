using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace Razor2Word.Utils
{
    public class WordHelpers
    {
        #region Constructors

        public WordHelpers(ViewContext viewContext,
            IViewDataContainer viewDataContainer)
            : this(viewContext, viewDataContainer, RouteTable.Routes)
        {
        }

        public WordHelpers(ViewContext viewContext,
            IViewDataContainer viewDataContainer, RouteCollection routeCollection)
        {
            ViewContext = viewContext;
            ViewData = new ViewDataDictionary(viewDataContainer.ViewData);
        }

        #endregion

        #region Properties

        public ViewDataDictionary ViewData { get; private set; }

        public ViewContext ViewContext { get; private set; }
        
        #endregion

        public StyleDescription Style(string id, FontFamily fontFamily = null, double? fontSize = null, FontStyle? fontStyle = null, Color? color = null)
        {
            var result = new StyleDescription(id, fontFamily, fontSize, fontStyle, color);
            ViewContext.Writer.Write(result.Definition());
            return result;
        }

        public DisposableHelper Table(ValueWithUnit? width = null, 
                                        Color? backColor = null, CellAligment? cellAligment = null, bool? useBorder = null)
        {
            if (width == null)
            {
                width = ValueWithUnit.HundreadPercent;
            }
            var tableAttributes = CreateTableAttributes(width, backColor, cellAligment, useBorder, null, null);
            var tag = string.Format("<table {0} >", tableAttributes);
            return CreateEnclosingTag(tag, "</table>");
        }

        public DisposableHelper TableRow(ValueWithUnit? height = null)
        {
            //TODO: pass style description here
            var heightAttr = string.Empty;
            if (height != null)
            {
                heightAttr = string.Format("height=\"{0}\" ", height);
            }
            return CreateEnclosingTag(string.Format("<tr {0}>", heightAttr), "</tr>");
        }

        public DisposableHelper TableCell(StyleDescription tableStyleDescription = null,
                                          ValueWithUnit? width = null,
                                          Color? backColor = null, CellAligment? cellAligment = null,
                                          int? rowSpan = null, int? colSpan = null)
        {
            var tableClass = (tableStyleDescription != null) ? tableStyleDescription.Reference().ToString() : string.Empty;
            var tableAttributes = CreateTableAttributes(width, backColor, cellAligment, null, rowSpan, colSpan);
            return CreateEnclosingTag(string.Format("<td {0}><p {1}>", tableAttributes, tableClass), "</p></td>");
        }

        public DisposableHelper Paragraph(StyleDescription styleDescription = null)
        {
            var style = (styleDescription != null) ? styleDescription.Reference().ToString() : string.Empty;
            return CreateEnclosingTag(string.Format("<p {0}>", style), "</p>");
        }

        [Obsolete("jeszcze nie zaimplementowane")] //TODO: implement me!!
        public DisposableHelper TextBlock(FontFamily fontFamily = null, double? fontSize = null, FontStyle? fontStyle = null, Color? color = null)
        {
            return CreateEnclosingTag(string.Format("<span>"), "</span>");
        } 

        #region Utils

        private DisposableHelper CreateEnclosingTag(string begin, string end)
        {
            return new DisposableHelper(() => ViewContext.Writer.Write(begin),
                                        () => ViewContext.Writer.Write(end));
        }

        private string CreateTableAttributes(ValueWithUnit? width, Color? backColor, CellAligment? cellAligment, bool? useBorder, int? rowSpan, int? colSpan)
        {
            string result = string.Empty;
            if (width != null)
            {
                result += string.Format("width=\"{0}\" ", width);
            }
            if (backColor != null)
            {
                result += "bgcolor=\"" + ColorTranslator.ToHtml(backColor.Value) + "\" ";
            }
            if (cellAligment != null)
            {
                if (cellAligment == CellAligment.Center)
                {
                    result += "align=\"center\" ";
                    result += "valign=\"middle\" ";
                }
                if ((cellAligment.Value & CellAligment.Left) == CellAligment.Left)
                {
                    result += "align=\"left\" ";
                    if (cellAligment == CellAligment.Left)
                    {
                        result += "valign=\"middle\" ";
                    }
                }
                if ((cellAligment.Value & CellAligment.Right) == CellAligment.Right)
                {
                    result += "align=\"right\" ";
                    if (cellAligment == CellAligment.Right)
                    {
                        result += "valign=\"middle\" ";
                    }
                }
                if ((cellAligment.Value & CellAligment.Top) == CellAligment.Top)
                {
                    if (cellAligment == CellAligment.Top)
                    {
                        result += "align=\"center\" ";
                    }
                    result += "valign=\"top\" ";
                }
                if ((cellAligment.Value & CellAligment.Bottom) == CellAligment.Bottom)
                {
                    if (cellAligment == CellAligment.Bottom)
                    {
                        result += "align=\"center\" ";
                    }
                    result += "valign=\"bottom\" ";
                }
            }
            if (useBorder == null || useBorder.Value)
            {
                result += "border=\"1\" ";
            }
            if (rowSpan != null)
            {
                result += "rowspan=\"" + rowSpan + "\" ";
            }
            if (colSpan != null)
            {
                result += "colspan=\"" + colSpan + "\" ";
            }

            return result;
        }

        #endregion

    }

    public class WordHelpers<TModel> : WordHelpers
    {
        public WordHelpers(ViewContext viewContext, IViewDataContainer container)
            : this(viewContext, container, RouteTable.Routes)
        {
        }

        public WordHelpers(ViewContext viewContext, IViewDataContainer container,
            RouteCollection routeCollection)
            : base(viewContext, container,
                routeCollection)
        {
            ViewData = new ViewDataDictionary<TModel>(container.ViewData);
        }

        public new ViewDataDictionary<TModel> ViewData { get; private set; }
    }

    public class DisposableHelper : IDisposable
    {
        private readonly Action end;

        public DisposableHelper(Action begin, Action end)
        {
            this.end = end;
            begin();
        }

        public void Dispose()
        {
            end();
        }
    }
}