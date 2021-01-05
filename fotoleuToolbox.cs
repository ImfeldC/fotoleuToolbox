using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using QRCoder;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;

namespace fotoleuToolbox
{
	public class fotoleuToolbox
	{
        private static Microsoft.Office.Tools.Excel.Worksheet s_debug_sheet = null;

        public fotoleuToolbox()
		{
		}

        public static void generateBill(string strFilePath)
        {
            Boolean bDebug = openDebugSheet();

            Microsoft.Office.Tools.Excel.Worksheet sheet = openFotoleuToolboxSheet("Auftragsblatt-Data");
            if( sheet != null )
            {
                try
                {
                    string pathTemplate = sheet.get_Range("G9").Value2.ToString();

                    generateDocument(pathTemplate, strFilePath, "");
                }
                catch (Exception ex)
                {
                    // Debug output
                    if (bDebug == true)
                    {
                        printDebugMessage("generateBill: Exception=" + ex.Message);
                    }
                    else
                    {
                        MessageBox.Show(ex.Message, "Bill Generator", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        public static void generateQRCode(string strFilePath)
        {
            Boolean bDebug = openDebugSheet();

            Microsoft.Office.Tools.Excel.Worksheet sheet = openFotoleuToolboxSheet("SwissQRCode");
            if (sheet != null)
            {
                try
                {
                    Microsoft.Office.Tools.Excel.Worksheet activesheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

                    string contactIBAN = sheet.get_Range("A2").Value2.ToString();
                    PayloadGenerator.SwissQrCode.Iban iban = new PayloadGenerator.SwissQrCode.Iban(contactIBAN, PayloadGenerator.SwissQrCode.Iban.IbanType.Iban);

                    string contactName = sheet.get_Range("A3").Value2.ToString();
                    string contactStreet = sheet.get_Range("A4").Value2.ToString();
                    string contactPlace = sheet.get_Range("A5").Value2.ToString();
                    string contactCountry = sheet.get_Range("A6").Value2.ToString();
                    //PayloadGenerator.SwissQrCode.Contact contact = new PayloadGenerator.SwissQrCode.Contact(contactName, "CH", contactStreet, contactPlace);
                    PayloadGenerator.SwissQrCode.Contact contact = PayloadGenerator.SwissQrCode.Contact.WithCombinedAddress(contactName, contactCountry, contactStreet, contactPlace);

                    string debitorName = sheet.get_Range("A12").Value2.ToString();
                    string debitorStreet = sheet.get_Range("A13").Value2.ToString();
                    string debitorPlace = sheet.get_Range("A14").Value2.ToString();
                    string debitorCountry = sheet.get_Range("A15").Value2.ToString();
                    //PayloadGenerator.SwissQrCode.Contact debitor = new PayloadGenerator.SwissQrCode.Contact(debitorName, "CH", debitorStreet, debitorPlace);
                    PayloadGenerator.SwissQrCode.Contact debitor = PayloadGenerator.SwissQrCode.Contact.WithCombinedAddress(debitorName, debitorCountry, debitorStreet, debitorPlace);

                    string additionalInfo1 = sheet.get_Range("A8").Value2.ToString();
                    string additionalInfo2 = sheet.get_Range("A9").Value2.ToString();
                    PayloadGenerator.SwissQrCode.AdditionalInformation additionalInformation = new PayloadGenerator.SwissQrCode.AdditionalInformation(additionalInfo1, additionalInfo2);

                    PayloadGenerator.SwissQrCode.Reference reference = new PayloadGenerator.SwissQrCode.Reference(PayloadGenerator.SwissQrCode.Reference.ReferenceType.NON);

                    decimal amount = (decimal)sheet.get_Range("A17").Value2;
                    //PayloadGenerator.SwissQrCode.Currency currency = PayloadGenerator.SwissQrCode.Currency.CHF;
                    PayloadGenerator.SwissQrCode.Currency currency;
                    if (sheet.get_Range("A18").Value2 == "CHF")
                    {
                        currency = PayloadGenerator.SwissQrCode.Currency.CHF;
                    }
                    else if (sheet.get_Range("A18").Value2 == "CHF")
                    {
                        currency = PayloadGenerator.SwissQrCode.Currency.EUR;
                    }
                    else
                    {
                        throw new Exception("Currency not supported: " + sheet.get_Range("A18").Value2);
                    }

                    PayloadGenerator.SwissQrCode generator = new PayloadGenerator.SwissQrCode(iban, currency, contact, reference, additionalInformation, debitor, amount);

                    QRCodeGenerator qrGenerator = new QRCodeGenerator();
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode(generator.ToString(), QRCodeGenerator.ECCLevel.M);
                    QRCode qrCode = new QRCode(qrCodeData);
                    Bitmap qrCodeAsBitmap = qrCode.GetGraphic(20, Color.Black, Color.White, Properties.Resources.CH_Kreuz_7mm, 14, 1);

                    // Temporary qrcode bitmap
                    string picturePath = Path.GetTempPath() + "qrcode.bmp";
                    if (File.Exists(picturePath))
                    {
                        File.Delete(picturePath);
                    }
                    qrCodeAsBitmap.Save(picturePath, ImageFormat.Bmp);

                    // alternative qrcode bitmap
                    string altpicturePath = sheet.get_Range("A26").Value2.ToString();
                    if (File.Exists(altpicturePath))
                    {
                        File.Delete(altpicturePath);
                    }
                    try
                    {   // save bitmap on alternative path
                        qrCodeAsBitmap.Save(altpicturePath, ImageFormat.Bmp);
                    }
                    catch
                    {
                        // catch expception, e.g. in case filepath is not valid/accesible
                        printDebugMessage("generateQRCode: Cannot save QR code bitmap to alternative path! altpicturePath=" + altpicturePath);
                    }

                    //sheet.Shapes.AddPicture(picturePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, 180, 40, 140, 140);
                    float Left = readFloatValue(sheet.get_Range("B21").Value2);
                    float Top = readFloatValue(sheet.get_Range("B22").Value2);
                    float Width = readFloatValue(sheet.get_Range("B23").Value2);
                    float Height = readFloatValue(sheet.get_Range("B24").Value2);
                    if (Left > 0)
                    {
                        sheet.Shapes.AddPicture(picturePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left, Top, Width, Height);
                    }

                    // Debug output
                    if (bDebug == true)
                    {
                        printDebugMessage("QR Code generated! Path=" + picturePath + " / AltPath=" + altpicturePath, contact.ToString(), debitor.ToString(), amount.ToString(), currency.ToString(), additionalInformation.UnstructureMessage, additionalInformation.BillInformation, iban.ToString());
                        printDebugImage(picturePath);
                    }

                    // Replace QR code bitmap in template
                    string strQRTemplatePath = sheet.get_Range("A29").Value2.ToString();
                    if (strFilePath.Equals(""))
                    {
                        if (sheet.get_Range("A31").Value2 != null)
                        {
                            strFilePath = sheet.get_Range("A31").Value2.ToString();
                        }
                    }
                    generateDocument(strQRTemplatePath, strFilePath, picturePath);

                    // delete temporary picture
                    File.Delete(picturePath);
                }
                catch (Exception ex)
                {
                    // Debug output
                    if (bDebug == true)
                    {
                        printDebugMessage("generateQRCode: Exception=" + ex.Message);
                    }
                    else
                    {
                        MessageBox.Show(ex.Message, "Swiss QR Code Generator", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        public static void generateDocument()
        {
            Boolean bDebug = openDebugSheet();
            string strAddDebugInfo = "";

            Microsoft.Office.Tools.Excel.Worksheet sheet = openFotoleuToolboxSheet("Auftragsblatt-Data");
            if (sheet != null)
            {
                try
                {
                    string pathTemplate = sheet.get_Range("G9").Value2.ToString();

                    string strFileName = readBookmarkValue(sheet, "Filename");
                    string strFilePath = readBookmarkValue(sheet, "Filepath");
                    string strFileTarget = strFilePath + strFileName; ;

                    string strAuftragID = readBookmarkValue(sheet, "AuftragID");
                    string strGUID = Guid.NewGuid().ToString();
                    // Temporary 1st word document
                    string strFile1 = Path.GetTempPath() + "1stDoc_" + strAuftragID + "_" + strGUID + ".docx";
                    if (File.Exists(strFile1))
                    {
                        File.Delete(strFile1);
                    }
                    // Temporary 2nd word document
                    string strFile2 = Path.GetTempPath() + "2ndDoc_" + strAuftragID + "_" + strGUID + ".docx";
                    if (File.Exists(strFile2))
                    {
                        File.Delete(strFile2);
                    }

                    generateBill(strFile1);     // generate billing information (w/o QR code)
                    generateQRCode(strFile2);   // generate QR code document
                    printDebugMessage("generateDocument: The two single files have been created! File1=" + strFile1 + ", File2=" + strFile2);

                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    strAddDebugInfo = "Word Application created!";

                    // Open first file and insert them
                    Microsoft.Office.Interop.Word.Document wordDoc1 = wordApp.Documents.Open(strFile1, ReadOnly: true);
                    wordDoc1.Fields.Update();
                    strAddDebugInfo = "Doc1 fields updated!";
                    wordDoc1.Activate();
                    strAddDebugInfo = "Doc1 activated!";
                    wordApp.Selection.WholeStory();
                    strAddDebugInfo = "Select 'WholeStory'!";
                    wordApp.Selection.Copy();
                    strAddDebugInfo = "Selection copied!";
                    // Open empty document
                    //Microsoft.Office.Interop.Word.Document wordDocTarget = wordApp.Documents.Open(strFileTarget);
                    Microsoft.Office.Interop.Word.Document wordDocTarget = wordApp.Documents.Add();
                    wordDocTarget.Activate();
                    strAddDebugInfo = "Target document activated (first time)!";
                    wordApp.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);
                    strAddDebugInfo = "clipboard into target document pasted!";
                    wordApp.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                    strAddDebugInfo = "Section break inserted into target document!";
                    wordDoc1.Close(SaveChanges: false);
                    strAddDebugInfo = "Doc1 document closed!";
                    wordDoc1 = null;
                    strAddDebugInfo = "First Document into traget document copied!";

                    // Open second file and insert them
                    Microsoft.Office.Interop.Word.Document wordDoc2 = wordApp.Documents.Open(strFile2, ReadOnly: true);
                    wordDoc2.Fields.Update();
                    wordDoc2.Activate();
                    wordApp.Selection.WholeStory();
                    wordApp.Selection.Copy();
                    wordDocTarget.Activate();
                    strAddDebugInfo = "Target document activated (second time)!";
                    wordApp.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);
                    //wordApp.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                    wordDoc2.Close(SaveChanges: false);
                    wordDoc2 = null;
                    strAddDebugInfo = "First Document into traget document copied!";

                    printDebugMessage("generateDocument: Target file has been created! Number of words=" + wordDocTarget.Words.Count);

                    #region Empty clipbord
                    // avoid word asking to keep clipboard when closing
                    // avoid message box: "do you want to keep last item you copied" at exit of word.
                    wordApp.Selection.ClearFormatting();    
                    wordApp.Selection.Find.ClearFormatting();
                    wordApp.Selection.Find.Replacement.ClearFormatting();
                    wordApp.Selection.InsertAfter(" ");  // select a "single" space
                    wordApp.Selection.Copy();            // "empty" clipboard, with a "single" space
                    #endregion

                    foreach (Microsoft.Office.Interop.Word.Section section in wordDocTarget.Sections)
                    {
                        // Do not link headers & footers with previous section
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = false;
                        section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                        section.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
                        section.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = false;
                    }

                    if (wordDocTarget.Sections.Count == 2)
                    {
                        // delete in 2nd section the header and footer
                        // NOTE: I had to use index 2; even I'm used to provide 0-based indexes. Not sure, but it works :-)
                        Section section = wordDocTarget.Sections[2];
                        //section.PageSetup.DifferentFirstPageHeaderFooter = -1; //=true (see also WdConstants.wdUndefined);
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete();
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Delete();
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Delete();
                        section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Delete();
                        section.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Delete();
                        section.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Delete();
                    }

                    wordApp.Visible = true;
                    wordApp.Activate();
                    wordDocTarget.SaveAs2(strFileTarget);
                    printDebugMessage("generateDocument: Target file has been saved! strFileTarget=" + strFileTarget);

                    wordDocTarget = null;
                    wordApp = null;

                }
                catch (Exception ex)
                {
                    // Debug output
                    if (bDebug == true)
                    {
                        printDebugMessage("generateDocument: Exception=" + ex.Message + ", strAddDebugInfo=" + strAddDebugInfo);
                    }
                    else
                    {
                        MessageBox.Show(ex.Message, "Document Generator", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private static void generateDocument(string pathTemplate, string pathFilename, string picturePath)
        {
            Boolean bDebug = openDebugSheet();

            if (File.Exists(pathTemplate))
            {
                string strFilename="";
                string strFilePath="";

                try
                {
                    Microsoft.Office.Tools.Excel.Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["Auftragsblatt-Data"]);
                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Open(pathTemplate, ReadOnly: true);

                    // Replace "bookmarks" within word document with real values from excel sheet
                    int replaceCounter = 0;
                    foreach (Microsoft.Office.Interop.Excel.ListObject table in sheet.ListObjects)
                    {
                        // The table "TabABBookmarks" contains three columns:
                        // 1st column: BookmarkName         -> name of the bookmark
                        // 2nd column: BookmarkValue        -> value which shall be insterted in final document
                        // 3rd column: BookmarksPlaceholder -> placeholder in template, which represents this bookmark; will be replaced with the value above. 
                        if (table.Name == "TabABBookmarks")
                        {
                            Microsoft.Office.Interop.Excel.Range tableRange = table.Range;

                            // Loop through rows ...
                            foreach (Microsoft.Office.Interop.Excel.Range row in tableRange.Rows)
                            {
                                string strBookmarkValue = "";
                                string strBookmarkPlaceholder = "";

                                // Get bookmark value (1. column)
                                string strBookmarkName = row.Cells[1, 1].Value2.ToString();
                                if (strBookmarkName.Equals("Filename"))
                                {
                                    // Get filename (from 2. column) to be used to save this document
                                    strFilename = row.Cells[1, 2].Value2.ToString();
                                }
                                if (strBookmarkName.Equals("Filepath"))
                                {
                                    // Get filepath (from 2. column) to be used to save this document
                                    strFilePath = row.Cells[1, 2].Value2.ToString();
                                }

                                // Get bookmark value (2. column)
                                strBookmarkValue = row.Cells[1, 2].Value2.ToString();

                                // Get bookmark value (3. column)
                                strBookmarkPlaceholder = row.Cells[1, 3].Value2.ToString();

                                // Replace bookmark ...
                                if (!strBookmarkValue.Equals(""))
                                {
                                    wordApp.Selection.Find.ClearFormatting();
                                    wordApp.Selection.Find.Replacement.ClearFormatting();

                                    wordApp.Selection.Find.Text = strBookmarkPlaceholder;
                                    wordApp.Selection.Find.Replacement.Text = strBookmarkValue;
                                    wordApp.Selection.Find.Forward = true;
                                    wordApp.Selection.Find.Wrap = WdFindWrap.wdFindAsk;
                                    wordApp.Selection.Find.Format = false;
                                    wordApp.Selection.Find.MatchCase = false;
                                    wordApp.Selection.Find.MatchWholeWord = false;
                                    wordApp.Selection.Find.MatchWildcards = false;
                                    wordApp.Selection.Find.MatchSoundsLike = false;
                                    wordApp.Selection.Find.MatchAllWordForms = false;

                                    bool bReplace = wordApp.Selection.Find.Execute(Replace: WdReplace.wdReplaceAll);
                                    if (bReplace)
                                    {
                                        replaceCounter++;
                                    }
                                }
                            }
                        }
                    }

                    // replace QR code bitmap with real bitmap
                    if(!picturePath.Equals(""))
                    {

                        // Replace QR image in shapes
                        // TODO: this code doesn't work yet; the image is NOT available/visible afterwards
                        foreach (Microsoft.Office.Interop.Word.Shape s in wordDoc.Shapes)
                        {
                            if (s.AlternativeText.ToUpper().Contains("QRCODE"))
                            {
                                // replace shape
                                Microsoft.Office.Interop.Word.Shape newShape = wordDoc.Shapes.AddPicture(picturePath, SaveWithDocument: true, Anchor: s.Anchor, Top: s.Top, Left: s.Left, Width: s.Width, Height: s.Height);
                                newShape.RelativeHorizontalPosition = s.RelativeHorizontalPosition;
                                newShape.RelativeHorizontalSize = s.RelativeHorizontalSize;
                                newShape.RelativeVerticalPosition = s.RelativeVerticalPosition;
                                newShape.RelativeVerticalSize = s.RelativeVerticalSize;
                                newShape.Top = s.Top;
                                newShape.Left = s.Left;
                                newShape.Width = s.Width;
                                newShape.Height = s.Height;
                                newShape.Visible = MsoTriState.msoTrue;

                                s.Delete();

                            }
                        }

                        // Replace QR image in "inline" shapes
                        foreach (Microsoft.Office.Interop.Word.InlineShape s in wordDoc.InlineShapes)
                        {
                            if (s.AlternativeText.ToUpper().Contains("QRCODE"))
                            {
                                Microsoft.Office.Interop.Word.Range range;

                                range = s.Range;
                                Microsoft.Office.Interop.Word.InlineShape newShape = wordDoc.InlineShapes.AddPicture(picturePath, SaveWithDocument: true, Range: range);
                                newShape.Width = s.Width;
                                newShape.Height = s.Height;

                                s.Delete();
                            }
                        }
                    }

                    wordDoc.Fields.Update();
                    wordDoc.Activate();

                    // save document
                    if(!strFilename.Equals(""))
                    {
                        wordDoc.SaveAs2(strFilePath + strFilename);
                    }

                    // save file OR show word app
                    if (!pathFilename.Equals(""))
                    {
                        wordDoc.SaveAs2(pathFilename);
                        wordDoc.Close(SaveChanges: false);
                        wordApp.Quit();
                    }
                    else
                    {
                        wordApp.Visible = true;
                        wordApp.Activate();
                    }


                    wordDoc = null;
                    wordApp = null;

                    // Debug output
                    if (bDebug == true)
                    {
                        printDebugMessage("generateDocument: Document generated! " + replaceCounter.ToString() + " bookmarks replaced. Template=" + pathTemplate + ", Filepath=" + pathFilename);
                    }
                }
                catch (Exception ex)
                {
                    // Debug output
                    if (bDebug == true)
                    {
                        printDebugMessage("generateDocument: Exception=" + ex.Message);
                    }
                    else
                    {
                        MessageBox.Show(ex.Message, "Document Generator", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                // Debug output
                if (bDebug == true)
                {
                    printDebugMessage("generateDocument: Document '" + pathTemplate + "' doesn't exists");
                }
            }
        }

        private static Microsoft.Office.Tools.Excel.Worksheet openFotoleuToolboxSheet( string strFotoleuSheetName)
        {
            Microsoft.Office.Tools.Excel.Worksheet sheet=null;

            try
            {
                sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[strFotoleuSheetName]);
            }
            catch (Exception)
            {
                MessageBox.Show("Please open a valid fotoleu excel workbook, which contains sheet '" + strFotoleuSheetName +"'", "fotoleu Toolbox", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return sheet;
        }

        private static float readFloatValue(dynamic value2)
        {
            try
            {
                if (value2 != null)
                {
                    return (float)value2;
                }
                else
                {
                    return 0;
                }
            }
            catch( Exception )
            {
                return 0;
            }
        }

        private static string readBookmarkValue(Microsoft.Office.Tools.Excel.Worksheet sheet, string strSearchBookmarkName)
        {
            string strBookmarkValueFound = "";

            // Replace "bookmarks" within word document with real values from excel sheet
            int bookmarkCounter = 0;
            foreach (Microsoft.Office.Interop.Excel.ListObject table in sheet.ListObjects)
            {
                // The table "TabABBookmarks" contains three columns:
                // 1st column: BookmarkName         -> name of the bookmark
                // 2nd column: BookmarkValue        -> value which shall be insterted in final document
                // 3rd column: BookmarksPlaceholder -> placeholder in template, which represents this bookmark; will be replaced with the value above. 
                if (table.Name == "TabABBookmarks")
                {
                    Microsoft.Office.Interop.Excel.Range tableRange = table.Range;

                    // Loop through rows ...
                    foreach (Microsoft.Office.Interop.Excel.Range row in tableRange.Rows)
                    {
                        // Get bookmark value (1. column)
                        string strBookmarkName = row.Cells[1, 1].Value2.ToString();
                        if (strBookmarkName.Equals(strSearchBookmarkName))
                        {
                            bookmarkCounter++;
                            // Get filename (from 2. column) to be used to save this document
                            strBookmarkValueFound = row.Cells[1, 2].Value2.ToString();
                            return strBookmarkValueFound;
                        }
                    }
                }
            }
            return strBookmarkValueFound;
        }

        private static bool openDebugSheet()
        {
            Boolean bDebug = false;
            try
            {
                if(s_debug_sheet == null)
                {
                    s_debug_sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["SwissQRCode-Debug"]);
                    if (s_debug_sheet.get_Range("D2").Value2 == "ON")
                    {
                        bDebug = true;
                    }
                }
                else
                {
                    // debug sheet has already been opened
                    bDebug = true;
                }
            }
            catch (Exception)
            {
                // Disable debugging; ignore exception
            }
            return bDebug;
        }

        private static void printDebugMessage(string strDebugMessage)
        {
            // Debug output
            if (s_debug_sheet != null)
            {
                Microsoft.Office.Interop.Excel.Range debugrows = s_debug_sheet.get_Range("A20");
                debugrows.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown); // sift down whole row

                Microsoft.Office.Interop.Excel.Range newdebugcell = null;
                newdebugcell = s_debug_sheet.get_Range("A20");
                newdebugcell.Value2 = DateTime.Now.ToString();
                newdebugcell = s_debug_sheet.get_Range("B20");
                newdebugcell.Value2 = strDebugMessage;
            }
        }

        private static void printDebugMessage(string strDebugMessage, string strContact, string strDebitor, string strAmount, string strCurrency, string strUnstructuredMessage, string strBillInfo, string strIBAN)
        {
            // Debug output
            if (s_debug_sheet != null)
            {
                Microsoft.Office.Interop.Excel.Range debugrows = s_debug_sheet.get_Range("A20");
                debugrows.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown); // sift down whole row

                Microsoft.Office.Interop.Excel.Range newdebugcell = null;
                newdebugcell = s_debug_sheet.get_Range("A20");
                newdebugcell.Value2 = DateTime.Now.ToString();
                newdebugcell = s_debug_sheet.get_Range("B20");
                newdebugcell.Value2 = strDebugMessage;
                newdebugcell = s_debug_sheet.get_Range("C20");
                newdebugcell.Value2 = "Contact: " + strContact;
                newdebugcell = s_debug_sheet.get_Range("D20");
                newdebugcell.Value2 = "Debitor: " + strDebitor;
                newdebugcell = s_debug_sheet.get_Range("E20");
                newdebugcell.Value2 = "Amount=" + strAmount + " / Currency=" + strCurrency;
                newdebugcell = s_debug_sheet.get_Range("F20");
                newdebugcell.Value2 = "Additional Information: UnstructureMessage=" + strUnstructuredMessage + " / BillInformation=" + strBillInfo;
                newdebugcell = s_debug_sheet.get_Range("G20");
                newdebugcell.Value2 = "IBAN: " + strIBAN;
            }
        }

        private static void printDebugImage(string strDebugImagePath)
        {
            // Debug output
            if (s_debug_sheet != null)
            {
                float debugLeft = readFloatValue(s_debug_sheet.get_Range("D5").Value2);
                float debugTop = readFloatValue(s_debug_sheet.get_Range("D6").Value2);
                float debugWidth = readFloatValue(s_debug_sheet.get_Range("D7").Value2);
                float debugHeight = readFloatValue(s_debug_sheet.get_Range("D8").Value2);
                if (debugLeft > 0)
                {
                    s_debug_sheet.Shapes.AddPicture(strDebugImagePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, debugLeft, debugTop, debugWidth, debugHeight);
                }
            }
        }
    }
}