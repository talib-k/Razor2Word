using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Windows;
using DocumentFormat.OpenXml.Wordprocessing;


namespace Razor2Word.Utils
{
    public struct PageDescription
    {
        public Size Size { get; set; }
        public PageMode Mode { get; set; }
        public PageMargins Margins { get; set; }

        public static PageDescription A4
        {
            get
            {
                return new PageDescription()
                {
                    Size = new Size() {Width = 210, Height = 297},
                    Mode = PageMode.Portrait,
                    Margins = new PageMargins() {Left = 25, Top = 25, Bottom = 25, Right = 25}
                };
            }
        }

        public PageDescription Landscape {
            get
            {
                return new PageDescription()
                {
                    Size = new Size() { Width = Size.Height, Height = Size.Width },
                    Mode = PageMode.Landscape,
                    Margins =
                        new PageMargins()
                        {
                            Left = Margins.Left,
                            Top = Margins.Top,
                            Bottom = Margins.Bottom,
                            Right = Margins.Right
                        }
                };
            }
        }

        public PageDescription Portrait
        {
            get
            {
                return new PageDescription()
                {
                    Size = new Size() { Width = Size.Height, Height = Size.Width },
                    Mode = PageMode.Portrait,
                    Margins =
                        new PageMargins()
                        {
                            Left = Margins.Left,
                            Top = Margins.Top,
                            Bottom = Margins.Bottom,
                            Right = Margins.Right
                        }
                };
            }
        }

        public PageDescription WithMargins(PageMargins margins)
        {
            return new PageDescription()
            {
                Size = new Size() { Width = Size.Width, Height = Size.Height },
                Mode = Mode,
                Margins =
                    new PageMargins()
                    {
                        Left = margins.Left,
                        Top = margins.Top,
                        Bottom = margins.Bottom,
                        Right = margins.Right
                    }
            };
        }

        public SectionProperties CreateSectionProperties()
        {
            SectionProperties result = new SectionProperties();
            var pageMargin = new PageMargin()
            {
                Top = MilisToTwips(Margins.Top),
                Bottom = MilisToTwips(Margins.Bottom),
                Left = (uint)MilisToTwips(Margins.Left),
                Right = (uint)MilisToTwips(Margins.Right)
            };

            var pageSize = new PageSize()
            {
                Width = (uint)MilisToTwips(Size.Width),
                Height = (uint)MilisToTwips(Size.Height),
                Orient = (Mode == PageMode.Portrait) ? PageOrientationValues.Portrait : PageOrientationValues.Landscape
            };
            result.Append(pageSize, pageMargin);

            return result;
        }

        private int MilisToTwips(double value)
        {
            return (int)(1440*value/25.4);
        }
    }

    public enum PageMode
    {
        Portrait,
        Landscape
    }

    public struct PageMargins
    {
        public PageMargins(double defaultMargin) : this()
        {
            Left = defaultMargin;
            Top = defaultMargin;
            Right = defaultMargin;
            Bottom = defaultMargin;
        }

        public double Left { get; set; }
        public double Top { get; set; }
        public double Right { get; set; }
        public double Bottom { get; set; }
    }
}