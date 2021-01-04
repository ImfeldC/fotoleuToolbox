using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using QRCoder;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace fotoleuToolbox
{
	public class fotoleuToolbox
	{
		public fotoleuToolbox()
		{
		}

        public void generateQRCodeInst()
		{
			generateQRCode();
		}

        public static void generateQRCode()
        {
            Boolean bDebug = false;
            Microsoft.Office.Tools.Excel.Worksheet debug_sheet = null;

            try
            {
                debug_sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["SwissQRCode-Debug"]);
                if (debug_sheet.get_Range("C2").Value2 == "ON")
                {
                    bDebug = true;
                }
            }
            catch (Exception)
            {
                // Disable debugging; ignore exception
            }

            try
            {
                Microsoft.Office.Tools.Excel.Worksheet activesheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
                Microsoft.Office.Tools.Excel.Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["SwissQRCode"]);

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
                qrCodeAsBitmap.Save(altpicturePath, ImageFormat.Bmp);

                //sheet.Shapes.AddPicture(picturePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, 180, 40, 140, 140);
                float Left = (float)sheet.get_Range("B21").Value2;
                float Top = (float)sheet.get_Range("B22").Value2;
                float Width = (float)sheet.get_Range("B23").Value2;
                float Height = (float)sheet.get_Range("B24").Value2;
                if (Left > 0)
                {
                    sheet.Shapes.AddPicture(picturePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left, Top, Width, Height);
                }

                // Debug output
                if (bDebug == true)
                {
                    Microsoft.Office.Interop.Excel.Range debugrows = debug_sheet.get_Range("A20");
                    debugrows.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown); // sift down whole row

                    Microsoft.Office.Interop.Excel.Range newdebugcell1 = debug_sheet.get_Range("A20");
                    newdebugcell1.Value2 = "QR Code generated! Path=" + picturePath + " / AltPath=" + altpicturePath + " at " + DateTime.Now.ToString();
                    Microsoft.Office.Interop.Excel.Range newdebugcell2 = debug_sheet.get_Range("B20");
                    newdebugcell2.Value2 = "Contact: " + contact.ToString();
                    Microsoft.Office.Interop.Excel.Range newdebugcell3 = debug_sheet.get_Range("C20");
                    newdebugcell3.Value2 = "Debitor: " + debitor.ToString();
                    Microsoft.Office.Interop.Excel.Range newdebugcell4 = debug_sheet.get_Range("D20");
                    newdebugcell4.Value2 = "Amount=" + amount.ToString() + " / Currency=" + currency.ToString();
                    Microsoft.Office.Interop.Excel.Range newdebugcell5 = debug_sheet.get_Range("E20");
                    newdebugcell5.Value2 = "Additional Information: UnstructureMessage=" + additionalInformation.UnstructureMessage + " / BillInformation=" + additionalInformation.BillInformation;
                    Microsoft.Office.Interop.Excel.Range newdebugcell6 = debug_sheet.get_Range("F20");
                    newdebugcell6.Value2 = "IBAN: " + iban.ToString();

                    float debugLeft = (float)debug_sheet.get_Range("C5").Value2;
                    float debugTop = (float)debug_sheet.get_Range("C6").Value2;
                    float debugWidth = (float)debug_sheet.get_Range("C7").Value2;
                    float debugHeight = (float)debug_sheet.get_Range("C8").Value2;
                    if (debugLeft > 0)
                    {
                        debug_sheet.Shapes.AddPicture(picturePath, MsoTriState.msoFalse, MsoTriState.msoCTrue, debugLeft, debugTop, debugWidth, debugHeight);
                    }
                }

                // Replace QR code bitmap in template
                string QRTemplatePath = sheet.get_Range("A29").Value2.ToString();
                generateDocument(QRTemplatePath, picturePath);

                // delete temporary picture
                File.Delete(picturePath);
            }
            catch (Exception ex)
            {
                // Debug output
                if (bDebug == true)
                {
                    Microsoft.Office.Interop.Excel.Range debugrows = debug_sheet.get_Range("A20");
                    debugrows.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown); // sift down whole row

                    Microsoft.Office.Interop.Excel.Range newdebugcell1 = debug_sheet.get_Range("A20");
                    newdebugcell1.Value2 = "generateQRCode: Exception=" + ex.Message + " at " + DateTime.Now.ToString();
                }
                else
                {
                    MessageBox.Show(ex.Message, "Swiss QR Code Generator", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        public static void generateBill()
        {
            Boolean bDebug = false;
            Microsoft.Office.Tools.Excel.Worksheet debug_sheet = null;

            try
            {
                debug_sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["SwissQRCode-Debug"]);
                if (debug_sheet.get_Range("C2").Value2 == "ON")
                {
                    bDebug = true;
                }
            }
            catch (Exception)
            {
                // Disable debugging; ignore exception
            }

            try
            {
                Microsoft.Office.Tools.Excel.Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["Auftragsblatt-Data"]);
                string pathTemplate = sheet.get_Range("G9").Value2.ToString();

                generateDocument(pathTemplate,"");
            }
            catch (Exception ex)
            {
                // Debug output
                if (bDebug == true)
                {
                    Microsoft.Office.Interop.Excel.Range debugrows = debug_sheet.get_Range("A20");
                    debugrows.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown); // sift down whole row

                    Microsoft.Office.Interop.Excel.Range newdebugcell1 = debug_sheet.get_Range("A20");
                    newdebugcell1.Value2 = "generateBill: Exception=" + ex.Message + " at " + DateTime.Now.ToString();
                }
                else
                {
                    MessageBox.Show(ex.Message, "Bill Generator", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        public static void generateDocument(string pathTemplate, string picturePath)
        {
            Boolean bDebug = false;
            Microsoft.Office.Tools.Excel.Worksheet debug_sheet = null;

            try
            {
                debug_sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["SwissQRCode-Debug"]);
                if (debug_sheet.get_Range("C2").Value2 == "ON")
                {
                    bDebug = true;
                }
            }
            catch (Exception)
            {
                // Disable debugging; ignore exception
            }

            if (File.Exists(pathTemplate))
            {
                try
                {
                    Microsoft.Office.Tools.Excel.Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["Auftragsblatt-Data"]);
                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Open(pathTemplate, ReadOnly: true);

                    // Replace "bookmarks" within word document with real values from excel sheet
                    int replaceCounter = 0;
                    foreach (Microsoft.Office.Interop.Excel.ListObject table in sheet.ListObjects)
                    {
                        if (table.Name == "TabABBookmarks")
                        {
                            Microsoft.Office.Interop.Excel.Range tableRange = table.Range;
                            String[] dataInRows = new string[tableRange.Rows.Count];

                            // Loop through rows ...
                            int i = 0;
                            foreach (Microsoft.Office.Interop.Excel.Range row in tableRange.Rows)
                            {
                                string strBookmarkValue = "";
                                string strBookmarkPlaceholder = "";

                                // For each row; loop through columns ...
                                for (int j = 0; j < row.Columns.Count; j++)
                                {
                                    dataInRows[i] = dataInRows[i] + "_" + row.Cells[1, j + 1].Value2.ToString();

                                    // Get bookmark value (2. column)
                                    if (j == 1)
                                    {
                                        strBookmarkValue = row.Cells[1, j + 1].Value2.ToString();
                                    }
                                    // Get bookmark value (3. column)
                                    if (j == 2)
                                    {
                                        strBookmarkPlaceholder = row.Cells[1, j + 1].Value2.ToString();
                                    }
                                }
                                i++;

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
                    if (!(picturePath == ""))
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

                    wordApp.Visible = true;
                    wordApp.Activate();

                    wordDoc = null;
                    wordApp = null;

                    // Debug output
                    if (bDebug == true)
                    {
                        Microsoft.Office.Interop.Excel.Range debugrows = debug_sheet.get_Range("A20");
                        debugrows.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown); // sift down whole row

                        Microsoft.Office.Interop.Excel.Range newdebugcell1 = debug_sheet.get_Range("A20");
                        newdebugcell1.Value2 = "generateDocument: Document generated! " + replaceCounter.ToString() + " bookmarks replaced. Path =" + pathTemplate + " at " + DateTime.Now.ToString();
                    }
                }
                catch (Exception ex)
                {
                    // Debug output
                    if (bDebug == true)
                    {
                        Microsoft.Office.Interop.Excel.Range debugrows = debug_sheet.get_Range("A20");
                        debugrows.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown); // sift down whole row

                        Microsoft.Office.Interop.Excel.Range newdebugcell1 = debug_sheet.get_Range("A20");
                        newdebugcell1.Value2 = "generateDocument: Exception=" + ex.Message + " at " + DateTime.Now.ToString();
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
                    Microsoft.Office.Interop.Excel.Range debugrows = debug_sheet.get_Range("A20");
                    debugrows.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown); // sift down whole row

                    Microsoft.Office.Interop.Excel.Range newdebugcell1 = debug_sheet.get_Range("A20");
                    newdebugcell1.Value2 = "generateDocument: Document '" + pathTemplate + "' doesn't exists; at " + DateTime.Now.ToString();
                }
            }
        }
    }
}