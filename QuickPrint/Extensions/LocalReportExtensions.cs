using Microsoft.Reporting.NETCore;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;

namespace QuickPrint.Extensions
{
    public static class LocalReportExtensions
    {
        public static void PrintToPrinter(this LocalReport report)
        {
            PageSettings pageSettings = new PageSettings();
            pageSettings.PaperSize = report.GetDefaultPageSettings().PaperSize;
            //pageSettings.Landscape = report.GetDefaultPageSettings().IsLandscape;
            pageSettings.Landscape = false;
            pageSettings.Margins = report.GetDefaultPageSettings().Margins;
            Print(report, pageSettings);
        }

        public static void Print(this LocalReport report, PageSettings pageSettings)
        {
            var height = pageSettings.PaperSize.Height;
            var width = pageSettings.PaperSize.Width;
            pageSettings.PaperSize.Height = width;
            pageSettings.PaperSize.Width = height;
            string deviceInfo =
                $@"<DeviceInfo>
                    <OutputFormat>EMF</OutputFormat>
                    <PageWidth>{pageSettings.PaperSize.Width}in</PageWidth>
                    <PageHeight>{pageSettings.PaperSize.Height}in</PageHeight>
                    <MarginTop>{pageSettings.Margins.Top}in</MarginTop>
                    <MarginLeft>{pageSettings.Margins.Left}in</MarginLeft>
                    <MarginRight>{pageSettings.Margins.Right}in</MarginRight>
                    <MarginBottom>{pageSettings.Margins.Bottom}in</MarginBottom>
                </DeviceInfo>";
            Warning[] warnings;
            var streams = new List<Stream>();
            var pageIndex = 0;
            report.Render("Image", deviceInfo,
                (name, fileNameExtension, encoding, mimeType, willSeek) =>
                {
                    MemoryStream stream = new MemoryStream();
                    streams.Add(stream);
                    return stream;
                }, out warnings);
            foreach (Stream stream in streams)
                stream.Position = 0;
            if (streams == null || streams.Count == 0)
                throw new Exception("لايوجد شيء لطباعته");
            using (var printDocument = new PrintDocument())
            {
                printDocument.DefaultPageSettings = pageSettings;
                if (!printDocument.PrinterSettings.IsValid)
                    throw new Exception("لايمكن إيجاد الطابعة الافتراضية");
                else
                {
                    printDocument.PrintPage += (sender, e) =>
                    {
                        Metafile pageImage = new Metafile(streams[pageIndex]);
                        Rectangle adjustedRect = new Rectangle(e.PageBounds.Left - (int)e.PageSettings.HardMarginX, e.PageBounds.Top - (int)e.PageSettings.HardMarginY, e.PageBounds.Width, e.PageBounds.Height);
                        e.Graphics.FillRectangle(Brushes.White, adjustedRect);
                        e.Graphics.DrawImage(pageImage, adjustedRect);
                        pageIndex++;
                        e.HasMorePages = (pageIndex < streams.Count);
                        e.Graphics.DrawRectangle(Pens.Red, adjustedRect);
                    };
                    printDocument.EndPrint += (Sender, e) =>
                    {
                        if (streams != null)
                        {
                            foreach (Stream stream in streams)
                                stream.Close();
                            streams = null;
                        }
                    };
                    printDocument.Print();
                }
            }
        }

        //public static void PrintReportToPrinter(this LocalReport report, PrinterSetting printerSetting)
        //{
        //    PageSettings pageSettings = new PageSettings();

        //    if (printerSetting != null)
        //    {
        //        pageSettings.PaperSize.Width = printerSetting.Width;
        //        pageSettings.PaperSize.Height = printerSetting.Height;
        //        pageSettings.Landscape = printerSetting.IsLandscape;
        //        pageSettings.Margins = new Margins
        //        {
        //            Top = printerSetting.MarginTop,
        //            Right = printerSetting.MarginRight,
        //            Bottom = printerSetting.MarginBottom,
        //            Left = printerSetting.MarginLeft,
        //        };
        //    }
        //    else
        //    {
        //        pageSettings.PaperSize = report.GetDefaultPageSettings().PaperSize;
        //        pageSettings.Landscape = report.GetDefaultPageSettings().IsLandscape;
        //        pageSettings.Margins = report.GetDefaultPageSettings().Margins;
        //    }
        //    Print(report, pageSettings);
        //}

        //private static List<Stream>? m_straems;
        //private static int m_currentPageIndex = 0;


        //public static void Print(this LocalReport report, PageSettings pageSettings)
        //{
        //    string deviceInfo =
        //        $@"<DeviceInfo>
        //            <OutputFormat>EMF</OutputFormat>
        //            <PageWidth>9.44882in</PageWidth>
        //            <PageHeight>6.29921in</PageHeight>
        //            <MarginTop>0in</MarginTop>
        //            <MarginLeft>0in</MarginLeft>
        //            <MarginRight>0in</MarginRight>
        //            <MarginBottom>0in</MarginBottom>
        //        </DeviceInfo>";
        //    Warning[] warnings;
        //    m_straems = new List<Stream>();
        //    report.Render("Image", deviceInfo, CreateStream, out warnings);
        //    foreach (Stream stream in m_straems)
        //        stream.Position = 0;

        //    if (m_straems == null || m_straems.Count == 0)
        //        throw new Exception("لايوجد شيء لطباعته");
        //    PrintDocument printDocument = new PrintDocument();
        //    if (!printDocument.PrinterSettings.IsValid)
        //    {
        //        throw new Exception("لايمكن إيجاد الطابعة الافتراضية");
        //    }
        //    else
        //    {
        //        printDocument.PrintPage += new PrintPageEventHandler(PrintPage);
        //        m_currentPageIndex = 0;
        //        printDocument.Print();
        //    }
        //}

        //private static Stream CreateStream(string name, string extension, Encoding encoding, string mimeType, bool willSeek)
        //{
        //    Stream stream = new MemoryStream();
        //    m_straems.Add(stream);
        //    return stream;
        //}

        //public static void PrintPage(object sender, PrintPageEventArgs e)
        //{
        //    Metafile pageImage = new Metafile(m_straems[m_currentPageIndex]);

        //    // Adjust rectangular area with printer margins.
        //    Rectangle adjustedRect = new Rectangle(
        //            e.PageBounds.Left - (int)e.PageSettings.HardMarginX,
        //            e.PageBounds.Top - (int)e.PageSettings.HardMarginY,
        //            e.PageBounds.Width,
        //            e.PageBounds.Height
        //        );

        //    // Draw a white background for the report
        //    e.Graphics.FillRectangle(Brushes.White, adjustedRect);

        //    // Draw the report content
        //    e.Graphics.DrawImage(pageImage, adjustedRect);

        //    // Prepare for next page. Make sure we haven't hit the end.
        //    m_currentPageIndex++;
        //    e.HasMorePages = m_currentPageIndex < m_straems.Count;

        //}

        //public static void DisposePrint()
        //{
        //    if (m_straems != null)
        //    {
        //        foreach (Stream stream in m_straems)
        //            stream.Close();
        //        m_straems = null;
        //    }
        //}



    }
}