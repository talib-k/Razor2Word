using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NotesFor.HtmlToOpenXml;

namespace Razor2Word.Utils
{
    public class WordGenerator
    {
        public WordGenerator(IView view)
        {
            View = view;
        }

        public IView View { get; private set; }
        public PageDescription PageConfiguration { get; set; }

        public byte[] GenerateWord(object model)
        {
            var html = CreateHtml(model);
            return Html2Word(html);
        }

        private string CreateHtml(object model)
        {
            var viewData = new ViewDataDictionary(model);
            using (var writer = new StringWriter())
            {
                var context = new ViewContext(new ControllerContext(new HttpContextWrapper(new HttpContext(new HttpRequest(string.Empty, "http://tempuri.org", string.Empty),
                                                                                                           new HttpResponse(new StringWriter()))),
                                                                    new RouteData(),
                                                                    new MockController()),
                                              View, viewData, new TempDataDictionary(), writer);
                View.Render(context, writer);
                writer.Flush();
                return writer.ToString();
            }
        }

        private byte[] Html2Word(string html)
        {
            using (var generatedDocument = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = package.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = package.AddMainDocumentPart();
                        new Document(new Body()).Save(mainPart);
                    }
                    ProcessStyles(html, mainPart);

                    var converter = new HtmlConverter(mainPart);

                    Body body = mainPart.Document.Body;

                    body.Append(PageConfiguration.CreateSectionProperties());

                    var paragraphs = converter.Parse(html);
                    foreach (OpenXmlCompositeElement paragraph in paragraphs)
                    {
                        body.Append(paragraph);
                    }

                    mainPart.Document.Save();
                }

                return generatedDocument.ToArray();
            }
        }

        private void ProcessStyles(string html, MainDocumentPart mainPart)
        {
            var styleDefinitionsPart = mainPart.StyleDefinitionsPart;

            if (styleDefinitionsPart == null)
            {
                styleDefinitionsPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                var root = new Styles();
                root.Save(styleDefinitionsPart);
                foreach (var styleDescription in StyleDescription.ReadFromHtml(html))
                {
                    AddNewStyle(styleDefinitionsPart, styleDescription, false);
                    AddNewStyle(styleDefinitionsPart, styleDescription, true);
                }
            }
        }

        private static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart, StyleDescription styleDescription, bool table)
        {
            const string normalStyleId = "Normalny";

            // Get access to the root element of the styles part.
            var styles = styleDefinitionsPart.Styles;

            Style style = null;
            StyleName styleName = null;
            if (table)
            {
                style = new Style()
                {
                    Type = StyleValues.Table,
                    StyleId = styleDescription.TableId,
                    CustomStyle = true
                };
                styleName = new StyleName() { Val = styleDescription.TableId };
            }
            else
            {
                style = new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = styleDescription.Id,
                    CustomStyle = true
                };
                styleName = new StyleName() { Val = styleDescription.Id };
            }

            style.Append(styleName);
            style.Append(new BasedOn() { Val = normalStyleId });
            style.Append(new NextParagraphStyle() { Val = normalStyleId });
            style.Append(new StyleParagraphProperties(new ContextualSpacing() { Val = false },
                new SpacingBetweenLines()
                {
                    Line = "240",
                    LineRule = LineSpacingRuleValues.Auto,
                    Before = "0",
                    After = "0"
                }));

            // Create the StyleRunProperties object and specify some of the run properties.
            var styleRunProperties = new StyleRunProperties();

            if ((styleDescription.FontStyle & FontStyle.Bold) == FontStyle.Bold)
            {
                styleRunProperties.Append(new Bold());
            }
            if ((styleDescription.FontStyle & FontStyle.Italic) == FontStyle.Italic)
            {
                styleRunProperties.Append(new Italic());
            }
            if ((styleDescription.FontStyle & FontStyle.Underline) == FontStyle.Underline)
            {
                styleRunProperties.Append(new Underline { Val = UnderlineValues.Single });
            }
            if (styleDescription.Color != null)
            {
                styleRunProperties.Append(new Color()
                {
                    Val = "#" +
                          styleDescription.Color.Value.R.ToString("X2") +
                          styleDescription.Color.Value.G.ToString("X2") +
                          styleDescription.Color.Value.B.ToString("X2")
                });
            }
            if (styleDescription.FontFamily != null)
            {
                styleRunProperties.Append(new RunFonts() { Ascii = styleDescription.FontFamily.Name });
            }
            if (styleDescription.FontSize != null)
            {
                styleRunProperties.Append(new FontSize() { Val = Convert.ToString(styleDescription.FontSize * 2) });
            }

            // Add the run properties to the style.
            style.Append(styleRunProperties);

            // Add the style to the styles part.
            styles.Append(style);
        }

        private class MockController : Controller
        {
        }
    }
}