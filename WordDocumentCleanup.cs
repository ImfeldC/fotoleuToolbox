using Microsoft.Office.Interop.Word;
using System;
using System.Linq;
using System.Xml.Linq;

namespace fotoleuToolbox
{
    public static class WordDocumentCleanup
    {
        public static void RemoveEmptyTrailingParagraphs(Document document, Action<string> log)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            if (log == null)
                throw new ArgumentNullException(nameof(log));
            if (document.TrackRevisions || document.Revisions.Count != 0)
            {
                log("Trailing-page cleanup skipped: document contains tracked changes.");
                return;
            }

            document.Repaginate();
            int pagesBefore = document.ComputeStatistics(WdStatistic.wdStatisticPages);
            int start = document.Content.End;
            int emptyParagraphs = 0;
            for (int i = document.Paragraphs.Count; i >= 1; i--)
            {
                Range range = document.Paragraphs[i].Range;
                if (!IsEmptyParagraph(document, range))
                    break;
                start = range.Start;
                emptyParagraphs++;
            }
            if (emptyParagraphs == 0)
            {
                log("Trailing-page cleanup: no unprotected empty trailing paragraphs; document unchanged.");
                return;
            }

            // Word requires the final paragraph mark, especially after a table.
            int end = document.Content.End - 1;
            if (start < end)
                document.Range(start, end).Delete();
            Range finalParagraph = document.Paragraphs.Last.Range;
            finalParagraph.Font.Size = 1;
            ParagraphFormat format = finalParagraph.ParagraphFormat;
            format.SpaceBefore = 0;
            format.SpaceAfter = 0;
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            format.LineSpacing = 1;
            format.PageBreakBefore = 0;
            format.KeepWithNext = 0;
            format.KeepTogether = 0;
            format.WidowControl = 0;
            document.Repaginate();
            int pagesAfter = document.ComputeStatistics(WdStatistic.wdStatisticPages);
            log("Trailing-page cleanup: removed " + (emptyParagraphs - 1) +
                " empty paragraph(s), normalized the required final paragraph; pages " +
                pagesBefore + " -> " + pagesAfter + ".");
        }

        private static bool IsEmptyParagraph(Document document, Range range)
        {
            string text = range.Text ?? "";
            if (text.Any(c => c != ' ' && c != '\t' && c != '\r' && c != '\n' &&
                c != '\v' && c != '\u00a0'))
                return false;
            if (range.Tables.Count != 0 || range.InlineShapes.Count != 0 ||
                range.Fields.Count != 0 || range.ContentControls.Count != 0 ||
                range.ListFormat.ListType != WdListType.wdListNoNumbering)
                return false;
            foreach (Border border in range.Borders)
                if (border.LineStyle != WdLineStyle.wdLineStyleNone)
                    return false;
            // WordOpenXML can omit a collapsed bookmark at a range boundary.
            Bookmarks bookmarks = document.Bookmarks;
            bool showHidden = bookmarks.ShowHidden;
            try
            {
                bookmarks.ShowHidden = true;
                foreach (Bookmark bookmark in bookmarks)
                    if (Overlaps(bookmark.Range, range))
                        return false;
            }
            finally
            {
                bookmarks.ShowHidden = showHidden;
            }
            foreach (ContentControl control in document.ContentControls)
                if (Overlaps(control.Range, range))
                    return false;
            foreach (Comment comment in document.Comments)
                if (Overlaps(comment.Scope, range))
                    return false;
            foreach (Shape shape in document.Shapes)
            {
                int anchor = shape.Anchor.Start;
                if (anchor >= range.Start && anchor < range.End)
                    return false;
            }

            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XElement body = XDocument.Parse(range.WordOpenXML).Descendants(word + "body").Single();
            XElement[] content = body.Elements().Where(element => element.Name != word + "sectPr").ToArray();
            if (content.Length != 1 || content[0].Name != word + "p")
                return false;
            XElement paragraph = content[0];
            XElement properties = paragraph.Element(word + "pPr");
            if (properties != null && properties.Elements().Any(element =>
                element.Name == word + "sectPr" || element.Name == word + "numPr" ||
                element.Name == word + "pBdr" || element.Name == word + "framePr" ||
                element.Name == word + "shd" || element.Name == word + "pPrChange"))
                return false;
            foreach (XElement element in paragraph.Elements())
            {
                if (element.Name == word + "pPr")
                    continue;
                if (element.Name != word + "r")
                    return false;
                foreach (XElement child in element.Elements())
                {
                    if (child.Name == word + "rPr" || child.Name == word + "t" ||
                        child.Name == word + "tab" || child.Name == word + "cr")
                        continue;
                    if (child.Name == word + "br" &&
                        ((string)child.Attribute(word + "type") ?? "textWrapping") == "textWrapping")
                        continue;
                    return false;
                }
            }
            return true;
        }

        private static bool Overlaps(Range anchor, Range range)
        {
            return anchor.StoryType == range.StoryType &&
                anchor.Start <= range.End && anchor.End >= range.Start;
        }
    }
}
