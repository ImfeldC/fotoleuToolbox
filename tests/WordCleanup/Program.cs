using fotoleuToolbox;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

internal static class Program
{
    private static int passed;

    [STAThread]
    private static int Main(string[] args)
    {
        Word.Application application = null;
        try
        {
            if (args.Length != 0 && (args.Length != 4 || args[0] != "--document" || args[2] != "--output"))
                throw new ArgumentException("Usage: --document <docx> --output <directory>");
            application = new Word.Application();
            application.Visible = false;
            application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            Test(application, "Whitespace-only final page is removed", document =>
            {
                document.Content.Text = "Invoice\r\r ";
                document.Paragraphs.Last.Format.PageBreakBefore = -1;
                document.Repaginate();
                Equal(2, Pages(document));
                Cleanup(document);
                Equal(1, Pages(document));
                Equal("Invoice", document.Content.Text.Trim());
            });
            Test(application, "Mandatory paragraph after a table is retained", document =>
            {
                document.Content.Text = "Invoice\r";
                Table table = document.Tables.Add(document.Range(document.Content.End - 1, document.Content.End - 1), 1, 1);
                table.Cell(1, 1).Range.Text = "Payment details";
                document.Paragraphs.Last.Range.InsertParagraphAfter();
                document.Paragraphs.Last.Format.PageBreakBefore = -1;
                Equal(2, Pages(document));
                string text = document.Content.Text.Trim();
                Cleanup(document);
                Equal(1, Pages(document));
                Equal(1, document.Tables.Count);
                Equal(text, document.Content.Text.Trim());
                Equal(1f, document.Paragraphs.Last.Range.Font.Size);
                Equal(1f, document.Paragraphs.Last.Format.LineSpacing);
            });
            Test(application, "Nonempty final paragraphs are unchanged", document =>
            {
                document.Content.Text = "Invoice\rKeep this text";
                string text = document.Content.Text;
                float size = document.Paragraphs.Last.Range.Font.Size;
                Cleanup(document);
                Equal(text, document.Content.Text);
                Equal(size, document.Paragraphs.Last.Range.Font.Size);
            });
            Test(application, "Inline QR image is preserved", document =>
            {
                string icon = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "..", "..", "..", "..", "Properties", "CH-Kreuz_7mm.png"));
                document.InlineShapes.AddPicture(icon, LinkToFile: false, SaveWithDocument: true,
                    Range: document.Range(0, 0));
                Cleanup(document);
                Equal(1, document.InlineShapes.Count);
            });
            Test(application, "Floating shape in an otherwise empty paragraph is preserved", document =>
            {
                Word.Shape shape = document.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                    20, 20, 200, 40, document.Range(0, 0));
                shape.TextFrame.TextRange.Text = "Keep this floating content";
                float size = document.Paragraphs.Last.Range.Font.Size;
                Cleanup(document);
                Equal(1, document.Shapes.Count);
                Equal(size, document.Paragraphs.Last.Range.Font.Size);
            });
            Test(application, "Bookmark on an empty final paragraph is preserved", document =>
            {
                document.Bookmarks.Add("KeepAtEnd", document.Range(0, 0));
                float size = document.Paragraphs.Last.Range.Font.Size;
                Cleanup(document);
                Equal(true, document.Bookmarks.Exists("KeepAtEnd"));
                Equal(size, document.Paragraphs.Last.Range.Font.Size);
            });
            Test(application, "Empty-result field is preserved", document =>
            {
                document.Fields.Add(document.Range(0, 0), WdFieldType.wdFieldEmpty, "QUOTE \"\"", false);
                Cleanup(document);
                Equal(1, document.Fields.Count);
            });
            Test(application, "Hidden bookmark on an empty final paragraph is preserved", document =>
            {
                document.Bookmarks.Add("_KeepAtEnd", document.Range(0, 0));
                bool showHidden = document.Bookmarks.ShowHidden;
                float size = document.Paragraphs.Last.Range.Font.Size;
                Cleanup(document);
                Equal(true, document.Bookmarks.Exists("_KeepAtEnd"));
                Equal(showHidden, document.Bookmarks.ShowHidden);
                Equal(size, document.Paragraphs.Last.Range.Font.Size);
            });
            Test(application, "Content control is preserved", document =>
            {
                document.ContentControls.Add(WdContentControlType.wdContentControlText, document.Range(0, 0));
                string text = document.Content.Text;
                Cleanup(document);
                Equal(1, document.ContentControls.Count);
                Equal(text, document.Content.Text);
            });
            Test(application, "Section break and page setup are preserved", document =>
            {
                document.Content.Text = "Invoice\r";
                document.Range(document.Content.End - 1, document.Content.End - 1)
                    .InsertBreak(WdBreakType.wdSectionBreakNextPage);
                document.Sections[2].PageSetup.LeftMargin = 60;
                Equal(2, document.Sections.Count);
                Cleanup(document);
                Equal(2, document.Sections.Count);
                Equal(60f, document.Sections[2].PageSetup.LeftMargin);
                Equal(2, Pages(document));
            });
            Test(application, "Visible paragraph border is preserved", document =>
            {
                document.Paragraphs.Last.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                float size = document.Paragraphs.Last.Range.Font.Size;
                Cleanup(document);
                Equal(WdLineStyle.wdLineStyleSingle,
                    document.Paragraphs.Last.Borders[WdBorderType.wdBorderBottom].LineStyle);
                Equal(size, document.Paragraphs.Last.Range.Font.Size);
            });
            Test(application, "Border inherited from a paragraph style is preserved", document =>
            {
                Style style = document.Styles.Add("PreservedBorder", WdStyleType.wdStyleTypeParagraph);
                style.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                document.Paragraphs.Last.Range.set_Style(style);
                float size = document.Paragraphs.Last.Range.Font.Size;
                Cleanup(document);
                Equal(WdLineStyle.wdLineStyleSingle,
                    document.Paragraphs.Last.Borders[WdBorderType.wdBorderBottom].LineStyle);
                Equal(size, document.Paragraphs.Last.Range.Font.Size);
            });
            Test(application, "Empty numbered paragraph is preserved", document =>
            {
                document.Paragraphs.Last.Range.ListFormat.ApplyBulletDefault();
                WdListType type = document.Paragraphs.Last.Range.ListFormat.ListType;
                float size = document.Paragraphs.Last.Range.Font.Size;
                Cleanup(document);
                Equal(type, document.Paragraphs.Last.Range.ListFormat.ListType);
                Equal(size, document.Paragraphs.Last.Range.Font.Size);
            });
            Test(application, "Tracked changes are not cleaned up", document =>
            {
                document.TrackRevisions = true;
                document.Content.Text = "Keep tracked text\r\r";
                string text = document.Content.Text;
                Cleanup(document);
                Equal(text, document.Content.Text);
                Equal(true, document.TrackRevisions);
            });
            if (args.Length != 0)
                SuppliedDocument(application, Path.GetFullPath(args[1]), Path.GetFullPath(args[3]));
            Console.WriteLine("Passed " + passed + " Word integration tests.");
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(exception);
            return 1;
        }
        finally
        {
            if (application != null)
                application.Quit(SaveChanges: WdSaveOptions.wdDoNotSaveChanges);
        }
    }

    private static void SuppliedDocument(Word.Application application, string path, string output)
    {
        Directory.CreateDirectory(output);
        Document document = application.Documents.Open(path, ReadOnly: true, AddToRecentFiles: false);
        try
        {
            string text = document.Content.Text.Trim();
            int[] counts = { document.Tables.Count, document.InlineShapes.Count, document.Shapes.Count, document.Sections.Count };
            int before = Pages(document);
            string name = Path.GetFileNameWithoutExtension(path);
            document.ExportAsFixedFormat(Path.Combine(output, name + "-before.pdf"), WdExportFormat.wdExportFormatPDF);
            Cleanup(document);
            Equal(text, document.Content.Text.Trim());
            int[] after = { document.Tables.Count, document.InlineShapes.Count, document.Shapes.Count, document.Sections.Count };
            if (!counts.SequenceEqual(after))
                throw new Exception("Cleanup changed the table, image or section counts.");
            int pages = Pages(document);
            if (pages > before)
                throw new Exception("Cleanup increased the page count.");
            document.SaveAs2(Path.Combine(output, name + "-cleaned.docx"));
            document.ExportAsFixedFormat(Path.Combine(output, name + "-after.pdf"), WdExportFormat.wdExportFormatPDF);
            Console.WriteLine("PASS Supplied document: pages " + before + " -> " + pages +
                "; body text, table/image counts and section count preserved.");
            passed++;
        }
        finally
        {
            document.Close(SaveChanges: WdSaveOptions.wdDoNotSaveChanges);
        }
    }

    private static void Test(Word.Application application, string name, Action<Document> test)
    {
        Document document = application.Documents.Add();
        try
        {
            test(document);
            Console.WriteLine("PASS " + name);
            passed++;
        }
        finally
        {
            document.Close(SaveChanges: WdSaveOptions.wdDoNotSaveChanges);
        }
    }

    private static int Pages(Document document)
    {
        document.Repaginate();
        return document.ComputeStatistics(WdStatistic.wdStatisticPages);
    }

    private static void Cleanup(Document document)
    {
        WordDocumentCleanup.RemoveEmptyTrailingParagraphs(document, Console.WriteLine);
    }

    private static void Equal<T>(T expected, T actual)
    {
        if (!Equals(expected, actual))
            throw new Exception("Expected " + expected + ", got " + actual + ".");
    }
}
