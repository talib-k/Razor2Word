using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Xml;
using DocumentFormat.OpenXml.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Office2013.PowerPoint.Roaming;

namespace Razor2Word.Utils
{
    public class StyleDescription
    {
        public string Id { get; private set; }
        public Color? Color { get; private set; }
        public FontStyle FontStyle { get; private set; }
        public double? FontSize { get; private set; }
        public FontFamily FontFamily { get; private set; }

        internal StyleDescription(string id, FontFamily fontFamily, double? fontSize, FontStyle? fontStyle, Color? color)
        {
            //TODO: validate id for null and spaces inside
            Id = id;
            FontFamily = fontFamily;
            FontSize = fontSize;
            Color = color;
            FontStyle = fontStyle ?? FontStyle.Regular;
        }

        internal static StyleDescription[] ReadFromHtml(string html)
        {
            var doc = new XmlDocument();
            doc.LoadXml(html);
            var result = new List<StyleDescription>();

            foreach (XmlElement element in doc.GetElementsByTagName("style"))
            {
                var tmp = element.InnerText.Split('{', '}');
                var id = tmp[0].Trim().Substring(1);
                var styleItems = tmp[1].Split(';')
                                        .Select(s => s.Trim())
                                        .Where(s => s.Length > 0)
                                        .ToDictionary(s => s.Substring(0, s.IndexOf(':')).Trim(),
                                            s => s.Substring(s.IndexOf(':') + 1).Trim());

                var color = (new ColorParser()).Read(styleItems);
                var fontStyle = (new FontParser()).Read(styleItems);
                var fontSize = (new FontSizeParser()).Read(styleItems);
                var fontFamily = (new FontFamilyParser()).Read(styleItems);
               
                result.Add(new StyleDescription(id, fontFamily, fontSize, fontStyle, color));
            }

            return result.ToArray();
        }

        public string TableId 
        {
            get { return Id + "__table"; }
        }

        public HtmlString Definition()
        {
            var result = new StringBuilder("<style type=\"text/css\"> ");
            result.AppendFormat(".{0} ", Id);
            result.Append("{ ");
            (new ColorParser()).Write(result, Color);
            (new FontParser()).Write(result, FontStyle);
            (new FontSizeParser()).Write(result, FontSize);
            (new FontFamilyParser()).Write(result, FontFamily);

            result.Append("} </style>");
            return new HtmlString(result.ToString());
        }

        public HtmlString Reference()
        {
            return new HtmlString(string.Format("class=\"{0}\"", Id));
        }

        public HtmlString TableReference()
        {
            return new HtmlString(string.Format("class=\"{0}\"", TableId));
        }

        protected interface IStyleItemParser<T>
        {
            T Read(Dictionary<string, string> styleItems);
            void Write(StringBuilder writer, T styleItem);
        }

        protected class FontParser : IStyleItemParser<FontStyle>
        {
            //TODO: line-through
            public FontStyle Read(Dictionary<string, string> styleItems)
            {
                var fontStyle = FontStyle.Regular;
                if (styleItems.ContainsKey("font-weight") && styleItems["font-weight"] == "bold")
                {
                    fontStyle |= FontStyle.Bold;
                }
                if (styleItems.ContainsKey("font-style") && styleItems["font-style"] == "italic")
                {
                    fontStyle |= FontStyle.Italic;
                }
                if (styleItems.ContainsKey("text-decoration") && styleItems["text-decoration"] == "underline")
                {
                    fontStyle |= FontStyle.Underline;
                }
                return fontStyle;
            }

            public void Write(StringBuilder writer, FontStyle styleItem)
            {
                if ((styleItem & FontStyle.Bold) == FontStyle.Bold)
                {
                    writer.Append("font-weight: bold; ");
                }
                if ((styleItem & FontStyle.Italic) == FontStyle.Italic)
                {
                    writer.Append("font-style: italic; ");
                }
                if ((styleItem & FontStyle.Underline) == FontStyle.Underline)
                {
                    writer.Append("text-decoration: underline; ");
                }
            }
        }

        protected class ColorParser : IStyleItemParser<Color?>
        {
            public Color? Read(Dictionary<string, string> styleItems)
            {
                Color? color = null;
                if (styleItems.ContainsKey("color"))
                {
                    color = ColorTranslator.FromHtml(styleItems["color"]);
                }
                return color;
            }

            public void Write(StringBuilder writer, Color? styleItem)
            {
                if (styleItem != null)
                {
                    writer.AppendFormat("color: {0}; ", ColorTranslator.ToHtml(styleItem.Value));
                }
            }
        }

        protected class FontSizeParser : IStyleItemParser<double?>
        {
            public double? Read(Dictionary<string, string> styleItems)
            {
                double? fontSize = null;
                if (styleItems.ContainsKey("font-size"))
                {
                    var fontSizeString = styleItems["font-size"];
                    if (fontSizeString.EndsWith("px"))
                    {
                        fontSize = Convert.ToDouble(styleItems["font-size"].Substring(0, fontSizeString.Length - 2), CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        throw new Exception(string.Format("Only px font sizes supported. Got {0}", fontSizeString));
                    }
                }
                return fontSize;
            }

            public void Write(StringBuilder writer, double? styleItem)
            {
                if (styleItem != null)
                {
                    writer.AppendFormat("font-size: {0}px; ", styleItem.Value.ToString(CultureInfo.InvariantCulture));
                }
            }
        }

        protected class FontFamilyParser : IStyleItemParser<FontFamily>
        {
            public FontFamily Read(Dictionary<string, string> styleItems)
            {
                FontFamily fontFamily = null;
                if (styleItems.ContainsKey("font-family"))
                {
                    fontFamily = FontFamily.FindFontFamily(styleItems["font-family"]);
                }
                return fontFamily;
            }

            public void Write(StringBuilder writer, FontFamily styleItem)
            {
                if (styleItem != null)
                {
                    writer.AppendFormat("font-family: {0}; ", styleItem.Name);
                }
            }
        }
    }

    public enum DimensionUnit
    {
        Percent,
        Point
    }

    public struct ValueWithUnit
    {
        public static ValueWithUnit HundreadPercent = new ValueWithUnit(100, DimensionUnit.Percent);

        private readonly int val;
        private readonly DimensionUnit unit;

        public int Value
        {
            get { return val; }
        }

        public DimensionUnit Unit
        {
            get { return unit; }
        }

        public static ValueWithUnit Percent(int value)
        {
            return new ValueWithUnit(value, DimensionUnit.Percent);
        }

        public static ValueWithUnit Point(int value)
        {
            return new ValueWithUnit(value, DimensionUnit.Point);
        }

        public ValueWithUnit(int value, DimensionUnit unit)
        {
            val = value;
            this.unit = unit;
        }

        private static string GetUnitMnemonic(DimensionUnit unit)
        {
            switch (unit)
            {
                case DimensionUnit.Percent:
                    return "%";
                case DimensionUnit.Point:
                    return "pt";
                default:
                    throw new ArgumentException("Unknown unit");
            }
        }

        public override string ToString()
        {
            return string.Format("{0}{1}", Value, GetUnitMnemonic(Unit));
        }
    }

    [Flags]
    public enum FontStyle
    {
        Regular = 0,
        Bold = 1,
        Italic = 2,
        Underline = 4
    }

    public class FontFamily
    {
        private static readonly Dictionary<string, FontFamily> AllFamilies = new Dictionary<string, FontFamily>();

        static FontFamily()
        {
            Arial = new FontFamily("Arial");
            AllFamilies.Add(Arial.Name, Arial);
            TimesNewRoman = new FontFamily("Times New Roman");
            AllFamilies.Add(TimesNewRoman.Name, TimesNewRoman);
        }

        internal static FontFamily FindFontFamily(string name)
        {
            return AllFamilies[name];
        }

        private FontFamily(string name)
        {
            Name = name;
        }

        public string Name { get; private set; }

        public static readonly FontFamily Arial;
        public static readonly FontFamily TimesNewRoman;
    }

    [Flags]
    public enum CellAligment
    {
        Center = 0,
        Top = 1,
        Left = 2,
        Right = 4,
        Bottom = 8,
        LeftTop = Left | Top,
        RightTop = Right | Top,
        LeftBottom = Left | Bottom,
        RightBottom = Right | Bottom
    }
}