using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using QRCoder;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using System.Deployment.Application;
using static QRCoder.PayloadGenerator.SwissQrCode.Reference;

namespace fotoleuToolbox
{
	public class fotoleuToolbox
	{
        private static string s_toolboxVersion = "";
        private static Microsoft.Office.Tools.Excel.Worksheet s_debug_sheet = null;
        private static Boolean s_bDebug = false;

        public fotoleuToolbox()
		{
		}

        public static void generateAuftragsblatt(string strFilePath)
        {
            Boolean bDebug = openDebugSheet();
            printDebugMessage("generateAuftragsblatt: Started .... (Toolbox Version: " + getCurrentToolboxVersion() + ")");

            string strVersion = readBookmarkValue("Version");
            if ( !strVersion.Equals("") )
            {
                printDebugMessage("generateAuftragsblatt: Start generating Auftragsblatt ... (strVersion=" + strVersion + ")");
                try
                {
                    string pathTemplate = readBookmarkValue("DocTemplate");

                    generateDocument(pathTemplate, strFilePath, "", false, "AB");
                }
                catch (Exception ex)
                {
                    // Debug output
                    if (bDebug == true)
                    {
                        printDebugMessage("generateAuftragsblatt: Exception=" + ex.Message);
                    }
                    else
                    {
                        MessageBox.Show(ex.Message, "Bill Generator", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// Generates QR bitmap code, replaces them in the template and stores the newly generated file under the name passed by strFilePath.
        /// </summary>
        /// <param name="strFilePath">Target filename for newly created document.</param>
        public static void generateQRCodeV2(string strFilePath)
        {
            Boolean bDebug = openDebugSheet();
            printDebugMessage("generateQRCode: Started .... (Toolbox Version: " + getCurrentToolboxVersion() + ")");

            string strVersion = readQRCodeValue("Version");
            if (!strVersion.Equals(""))
            {
                printDebugMessage("generateQRCode: Start generating QR codes ... (strVersion=" + strVersion + ")");
                try
                {
                    #region Read Values from table and create objects
                    string contactIBAN = readQRCodeValue("IBAN");
                    PayloadGenerator.SwissQrCode.Iban iban = new PayloadGenerator.SwissQrCode.Iban(contactIBAN, PayloadGenerator.SwissQrCode.Iban.IbanType.Iban);
                    printDebugMessage("generateQRCode: Value read from table! contactIBAN=" + contactIBAN);

                    string contactName = readQRCodeValue("ContactName");
                    string contactStreet = readQRCodeValue("ContactAdressLine1");
                    string contactPlace = readQRCodeValue("ContactAdressLine2");
                    string contactCountry = readQRCodeValue("ContactCountry");
                    printDebugMessage("generateQRCode: Contact Values read from table! contactName=" + contactName);
                    //PayloadGenerator.SwissQrCode.Contact contact = new PayloadGenerator.SwissQrCode.Contact(contactName, "CH", contactStreet, contactPlace);
                    PayloadGenerator.SwissQrCode.Contact contact = PayloadGenerator.SwissQrCode.Contact.WithCombinedAddress(contactName, contactCountry, contactStreet, contactPlace);

                    string debitorName = readQRCodeValue("DebitorName");
                    string debitorStreet = readQRCodeValue("DebitorAdressLine1");
                    string debitorPlace = readQRCodeValue("DebitorAdressLine2");
                    string debitorCountry = readQRCodeValue("DebitorCountry");
                    printDebugMessage("generateQRCode: Debitor Values read from table! debitorName=" + debitorName);
                    //PayloadGenerator.SwissQrCode.Contact debitor = new PayloadGenerator.SwissQrCode.Contact(debitorName, "CH", debitorStreet, debitorPlace);
                    PayloadGenerator.SwissQrCode.Contact debitor = PayloadGenerator.SwissQrCode.Contact.WithCombinedAddress(debitorName, debitorCountry, debitorStreet, debitorPlace);

                    string additionalInfo1 = readQRCodeValue("UnstructureMessage");
                    string additionalInfo2 = readQRCodeValue("BillInformation");
                    PayloadGenerator.SwissQrCode.AdditionalInformation additionalInformation = new PayloadGenerator.SwissQrCode.AdditionalInformation(additionalInfo1, additionalInfo2);

                    PayloadGenerator.SwissQrCode.Reference reference;
                    string strReference = readQRCodeValue("SCOR");
                    if (!strReference.Equals(""))
                    {
                        reference = new PayloadGenerator.SwissQrCode.Reference(PayloadGenerator.SwissQrCode.Reference.ReferenceType.SCOR, strReference, ReferenceTextType.CreditorReferenceIso11649);
                    }
                    else
                    {
                        reference = new PayloadGenerator.SwissQrCode.Reference(PayloadGenerator.SwissQrCode.Reference.ReferenceType.NON);
                    }
                    printDebugMessage("generateQRCode: Values read from table! strReference=" + strReference);
                    #endregion

                    string strAmount = readQRCodeValue("Amount");
                    decimal amount = -1;
                    try
                    {
                        amount = decimal.Parse(strAmount);
                    }
                    catch
                    {
                        // value cannot be converted in a "valid" decimal --> skip QR code production
                        printDebugMessage("generateQRCode: Amount value cannot be converted in a 'valid' decimal --> skip QR code production! strAmount=" + strAmount);
                        amount = -1;
                    }
                    if (amount >= 0)
                    {
                        #region Retrieve Currency
                        //PayloadGenerator.SwissQrCode.Currency currency = PayloadGenerator.SwissQrCode.Currency.CHF;
                        string strCurrency = readQRCodeValue("Currency");
                        PayloadGenerator.SwissQrCode.Currency currency;
                        if (strCurrency.Equals("CHF"))
                        {
                            currency = PayloadGenerator.SwissQrCode.Currency.CHF;
                        }
                        else if (strCurrency.Equals("EUR"))
                        {
                            currency = PayloadGenerator.SwissQrCode.Currency.EUR;
                        }
                        else
                        {
                            throw new Exception("Currency not supported: strCurrency=" + strCurrency);
                        }
                        #endregion

                        // Create QR Code
                        PayloadGenerator.SwissQrCode generator = new PayloadGenerator.SwissQrCode(iban, currency, contact, reference, additionalInformation, debitor, amount);

                        QRCodeGenerator qrGenerator = new QRCodeGenerator();
                        QRCodeData qrCodeData = qrGenerator.CreateQrCode(generator.ToString(), QRCodeGenerator.ECCLevel.M);
                        QRCode qrCode = new QRCode(qrCodeData);
                        Bitmap qrCodeAsBitmap = qrCode.GetGraphic(20, Color.Black, Color.White, Properties.Resources.CH_Kreuz_7mm, 14, 1);

                        #region Save QR code bitmap, to be used in further processing
                        // Temporary qrcode bitmap
                        string picturePath = Path.GetTempPath() + "qrcode.bmp";
                        if (File.Exists(picturePath))
                        {
                            File.Delete(picturePath);
                        }
                        qrCodeAsBitmap.Save(picturePath, ImageFormat.Bmp);
                        printDebugMessage("generateQRCode: QR code bitmap saved! picturePath=" + picturePath);
                        #endregion

                        #region Save QR code bitmap to an addtional (atlernative) file
                        // alternative qrcode bitmap
                        string altpicturePath = readQRCodeValue("BitmapPath");
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
                        printDebugMessage("generateQRCode: Alternative QR code bitmap saved! altpicturePath=" + altpicturePath);
                        #endregion

                        // Debug output
                        if (bDebug == true)
                        {
                            printDebugMessage("QR Code generated! Path=" + picturePath + " / AltPath=" + altpicturePath, contact.ToString(), debitor.ToString(), amount.ToString(), currency.ToString(), additionalInformation.UnstructureMessage, additionalInformation.BillInformation, iban.ToString());
                            printDebugImage(picturePath);
                        }

                        // Replace QR code bitmap in template
                        if (strFilePath.Equals(""))
                        {
                            strFilePath = readQRCodeValue("QRFilePath");
                        }
                        string strQRTemplatePath = readQRCodeValue("QRTemplate");
                        generateDocument(strQRTemplatePath, strFilePath, picturePath, false, "ORESZ");

                        // delete temporary picture
                        File.Delete(picturePath);
                    }

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
            else
            {
                MessageBox.Show("Please open a valid fotoleu excel workbook, which contains OR code value table!", "Swiss QR Code Generator", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public static void generateRechnung()
        {
            Microsoft.Office.Interop.Word.Application wordApp = null;

            Boolean bDebug = openDebugSheet();
            string strAddDebugInfo = "";
            printDebugMessage("generateRechnung: Started .... (Toolbox Version: " + getCurrentToolboxVersion() + ")");

            string strVersion = readBookmarkValue("Version");
            if (!strVersion.Equals(""))
            {
                printDebugMessage("generateRechnung: Start generating bill ... (strVersion=" + strVersion + ")");
                try
                {
                    string pathTemplate = readBookmarkValue("DocTemplate");
                    strAddDebugInfo = "Template path read! pathTemplate=" + pathTemplate;

                    string strAuftragID = readBookmarkValue("AuftragID");
                    strAddDebugInfo = "AuftragID read! strAuftragID=" + strAuftragID;
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

                    generateAuftragsblatt(strFile1);     // generate billing information (w/o QR code)
                    #region Check that file exists
                    if ( File.Exists(strFile1))
                    {
                        printDebugMessage("generateRechnung: The first file has been created! File1=" + strFile1);
                    }
                    else
                    {
                        printDebugMessage("generateRechnung: The first file has NOT been created! File1=" + strFile1);
                    }
                    #endregion
                    generateQRCodeV2(strFile2);   // generate QR code document
                    #region Check that file exists
                    if (File.Exists(strFile2))
                    {
                        printDebugMessage("generateRechnung: The second file has been created! File2=" + strFile2);
                    }
                    else
                    {
                        printDebugMessage("generateRechnung: The second file has NOT been created! File2=" + strFile2);
                    }
                    #endregion

                    if (File.Exists(strFile1) && File.Exists(strFile2))
                    {
                        wordApp = new Microsoft.Office.Interop.Word.Application();
                        strAddDebugInfo = "Word Application created!";

                        #region Open files and copy them to target file.
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
                        #endregion

                        printDebugMessage("generateRechnung: Target file has been created! Number of words=" + wordDocTarget.Words.Count);

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

                        string strFileName = readBookmarkValue("Filename");
                        string strFilePath = readBookmarkValue("Filepath");
                        string strFileTarget = strFilePath + strFileName; ;
                        strAddDebugInfo = "Target filename and path read! strFileName=" + strFileName + ", strFilePath=" + strFilePath + ", strFileTarget=" + strFileTarget;

                        if (Directory.Exists(strFilePath))
                        {
                            wordDocTarget.SaveAs2(strFileTarget);
                            printDebugMessage("generateRechnung: Target file has been saved! strFileTarget=" + strFileTarget);
                        }
                        else
                        {
                            printDebugMessage("generateRechnung: Path doesn't exists, cannot save file! strFilePath=" + strFilePath + ", strFileName=" + strFileName);
                        }
                        wordApp.Visible = true;
                        wordApp.Activate();

                        wordDocTarget = null;
                        wordApp = null;

                    }
                    else
                    {
                        // One or both files doesn't exists -> skip production of combined document
                        printDebugMessage("generateRechnung: One or both files doesn't exist -> skip production of combined document");
                    }

                    #region Delete temporary files
                    if (!bDebug)
                    {
                        // Delete temporary files
                        if (File.Exists(strFile1))
                        {
                            File.Delete(strFile1);
                        }
                        if (File.Exists(strFile2))
                        {
                            File.Delete(strFile2);
                        }
                        printDebugMessage("generateRechnung: The two single files have been deleted! File1=" + strFile1 + ", File2=" + strFile2);
                    }
                    #endregion

                }
                catch (Exception ex)
                {
                    // Debug output
                    if (bDebug == true)
                    {
                        printDebugMessage("generateRechnung: Exception=" + ex.Message + ", strAddDebugInfo=" + strAddDebugInfo);
                    }
                    else
                    {
                        MessageBox.Show(ex.Message, "Document Generator", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    if( (wordApp!=null) && (!wordApp.Visible) )
                    {
                        // Word APP is not visible yet -> close it
                        wordApp.Quit(SaveChanges: false);
                    }
                }
            }
        }

        /// <summary>
        /// Private method to generate document.
        /// Replaces the bookmarks in the file template with real values read from excel sheet.
        /// </summary>
        /// <param name="pathTemplate">Name and path of the template.</param>
        /// <param name="pathFilename">Name and path of the target document.</param>
        /// <param name="picturePath">Name and path of the QR code bitmap to be used in target document.</param>
        /// <param name="bSaveDocument">Create a copy of target docment, using filename and filepath read from excel sheet.</param>
        /// <param name="strFileNameSuffix">Suffix to be added to filename read from excel sheet.</param>
        private static void generateDocument(string pathTemplate, string pathFilename, string picturePath, bool bSaveDocument, string strFileNameSuffix)
        {
            Boolean bDebug = openDebugSheet();

            if (File.Exists(pathTemplate))
            {
                Microsoft.Office.Interop.Word.Application wordApp = null;
                try
                {
                    wordApp = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Open(pathTemplate, ReadOnly: true);
                    int replaceCounter = 0;

                    foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
                    {
                        // Replace "bookmarks" within word document with real values from excel sheet
                        foreach (Microsoft.Office.Interop.Excel.ListObject table in worksheet.ListObjects)
                        {
                            // The table "TabABBookmarks" contains three columns:
                            // 1st column: BookmarkName         -> name of the bookmark
                            // 2nd column: BookmarkValue        -> value which shall be insterted in final document
                            // 3rd column: BookmarksPlaceholder -> placeholder in template, which represents this bookmark; will be replaced with the value above. 
                            if (table.Name == "TabABBookmarks")
                            {
                                #region Replace "bookmarks" within word document with real values from excel sheet
                                Microsoft.Office.Interop.Excel.Range tableRange = table.Range;

                                // Loop through rows ...
                                foreach (Microsoft.Office.Interop.Excel.Range row in tableRange.Rows)
                                {
                                    string strBookmarkValue = "";
                                    string strBookmarkPlaceholder = "";

                                    // Get bookmark value (1. column)
                                    string strBookmarkName = row.Cells[1, 1].Value2.ToString();

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
                                #endregion
                            }
                        }
                    }

                    #region replace QR code bitmap with real bitmap
                    // replace QR code bitmap with real bitmap
                    if (!picturePath.Equals(""))
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
                    #endregion

                    wordDoc.Fields.Update();
                    wordDoc.Activate();

                    #region save document with filename and path read from excel sheet
                    // save document with filename and path read from excel sheet
                    if (bSaveDocument)
                    {
                        string strFileName = readBookmarkValue("Filename");
                        if (!strFileName.Equals(""))
                        {
                            string strFilePath = readBookmarkValue("Filepath");
                            if (Directory.Exists(strFilePath))
                            {
                                // add suffix - if available - to the filename
                                if (!strFileNameSuffix.Equals(""))
                                {
                                    strFileName = strFileName.Replace(".docx", "_" + strFileNameSuffix + ".docx");
                                }
                                wordDoc.SaveAs2(strFilePath + strFileName);
                                printDebugMessage("generateDocument: Document saved! strFileName=" + strFileName + ", strFilePath=" + strFilePath);
                            }
                            else
                            {
                                printDebugMessage("generateDocument: Path doesn't exists, cannot save file! strFileName=" + strFileName + ", strFilePath=" + strFilePath);
                            }
                        }
                    }
                    #endregion

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

                    printDebugMessage("generateDocument: Document generated! " + replaceCounter.ToString() + " bookmarks replaced. Template=" + pathTemplate + ", Filepath=" + pathFilename);
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
                    if ((wordApp != null) && (!wordApp.Visible))
                    {
                        // Word APP is not visible yet -> close it
                        wordApp.Quit(SaveChanges: false);
                    }
                }
            }
            else
            {
                printDebugMessage("generateDocument: Document '" + pathTemplate + "' doesn't exists");
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

        /// <summary>
        /// Reads the value for a named attribute passed by strValueName from a table with the name passed by strTablename.
        /// The table contains three columns:
        /// 1st column: Name         -> name of the attribute
        /// 2nd column: Value        -> value which shall be insterted in final document
        /// 3rd column: Placeholder  -> placeholder in template, which represents this attribute; will be replaced with the value above. 
        /// </summary>
        /// <param name="strTablename">Table name.</param>
        /// <param name="ValueName">Value name.</param>
        /// <returns></returns>
        private static string readValueFromTable(string strTablename, string strValueName)
        {
            string strValueFound = "";

            foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                foreach (Microsoft.Office.Interop.Excel.ListObject table in worksheet.ListObjects)
                {
                    if (table.Name.Equals(strTablename))
                    {
                        Microsoft.Office.Interop.Excel.Range tableRange = table.Range;

                        // Loop through rows ...
                        foreach (Microsoft.Office.Interop.Excel.Range row in tableRange.Rows)
                        {
                            // Get attribute name (1. column)
                            if(row.Cells[1, 1].Value2 != null)
                            {
                                string strName = row.Cells[1, 1].Value2.ToString();
                                if (strName.Equals(strValueName))
                                {
                                    // Get attribute value (from 2. column) to be used to save this document
                                    if (row.Cells[1, 2].Value2 != null)
                                    {
                                        strValueFound = row.Cells[1, 2].Value2.ToString();
                                        return strValueFound;
                                    }
                                    else
                                    {
                                        return "";
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return strValueFound;
        }

        /// <summary>
        /// Reads a bookmark value from a table with name "TabABBookmarks" in activate workbook. It retruns the bookmark value for the bookmark with the name passed in by strSearchBookmarkName.
        /// </summary>
        /// <param name="strSearchBookmarkName">Name of the bookmark to be returned.</param>
        /// <returns>Value of the bookmark. Returns empty string if bookmark hasn't been found.</returns>
        private static string readBookmarkValue(string strSearchBookmarkName)
        {
            return readValueFromTable("TabABBookmarks", strSearchBookmarkName);
        }

        /// <summary>
        /// Reads a value for a named attribute from table "TabQRCode" in active worksheets.
        /// </summary>
        /// <param name="strQRCodeAttributeName">Name of the attribute.</param>
        /// <returns>Value of the attribute. Returns "" if not found.</returns>
        private static string readQRCodeValue(string strQRCodeAttributeName)
        {
            return readValueFromTable("TabQRCode", strQRCodeAttributeName);
        }

        private static bool openDebugSheet()
        {
            try
            {
                if(s_debug_sheet == null)
                {
                    s_debug_sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["SwissQRCode-Debug"]);
                }
                if (s_debug_sheet.get_Range("D2").Value2 == "ON")
                {
                    s_bDebug = true;
                }
                else
                {
                    s_bDebug = false;
                }
            }
            catch (Exception)
            {
                // Disable debugging; ignore exception
                s_debug_sheet = null;
                s_bDebug = false;
            }
            return s_bDebug;
        }

        private static void printDebugMessage(string strDebugMessage)
        {
            // Debug output
            if ((s_debug_sheet != null) && s_bDebug)
            {
                Microsoft.Office.Interop.Excel.Range debugrows = s_debug_sheet.get_Range("A20");
                debugrows.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown); // sift down whole row

                Microsoft.Office.Interop.Excel.Range newdebugcell = null;
                newdebugcell = s_debug_sheet.get_Range("A20");
                newdebugcell.Value2 = DateTime.Now.ToString();
                newdebugcell = s_debug_sheet.get_Range("B20");
                newdebugcell.Value2 = strDebugMessage;
            }

            /* doesn't work :-(
            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.DisplayStatusBar = true;
                excelApp.StatusBar = "Debug Message: " + strDebugMessage;
            }
            catch
            {
                // Ignore any exception
                string strStatusBar = "Debug Message: " + strDebugMessage;
            }
            */
        }

        private static void printDebugMessage(string strDebugMessage, string strContact, string strDebitor, string strAmount, string strCurrency, string strUnstructuredMessage, string strBillInfo, string strIBAN)
        {
            // Debug output
            if ((s_debug_sheet != null) && s_bDebug)
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
            if ((s_debug_sheet != null) && s_bDebug)
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

        public static string getCurrentToolboxVersion()
        {
            if( s_toolboxVersion.Equals(""))
            {
                try
                {
                    System.Version ver = ApplicationDeployment.CurrentDeployment.CurrentVersion;
                    s_toolboxVersion = ver.Major.ToString() + "." + ver.Minor.ToString() + "." + ver.Build.ToString() + "." + ver.Revision.ToString();
                }
                catch
                {
                    s_toolboxVersion = "n/a";
                }
            }
            return s_toolboxVersion;
        }
    }
}