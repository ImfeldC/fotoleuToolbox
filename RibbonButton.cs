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

namespace SwissQRCode
{
    public partial class RibbonButton
    {
        private void RibbonButton_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void buttonGenerate_Click(object sender, RibbonControlEventArgs e)
        {
            SwissQRCode.generateQRCode();
        }

        private void buttonDocument_Click(object sender, RibbonControlEventArgs e)
        {
            SwissQRCode.generateDocument();
        }
    }
}
