using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
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
    public partial class RibbonButton
    {
        private void RibbonButton_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnAuftragsblatt_Click(object sender, RibbonControlEventArgs e)
        {
            fotoleuToolbox.printDebugMessage("----- Create Auftragsblatt -----", "btnAuftragsblatt_Click");
            fotoleuToolbox.generateAuftragsblatt("",true);
        }

        private void btnQR_Click(object sender, RibbonControlEventArgs e)
        {
            fotoleuToolbox.printDebugMessage("----- Create QR Code -----", "btnQR_Click");
            fotoleuToolbox.generateQRCodeV2("",true);
        }

        private void btnRechnung_Click(object sender, RibbonControlEventArgs e)
        {
            fotoleuToolbox.printDebugMessage("----- Create Rechnung -----", "btnRechnung_Click");
            fotoleuToolbox.generateRechnung();
        }

    }
}
