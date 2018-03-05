using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using ZXing;

namespace BioBank
{
    public class printFunction
    {

        public static bool IsPrinterExist(string mPrinterName)
        {
            int i = 0;
            for (i = 0; i <= System.Drawing.Printing.PrinterSettings.InstalledPrinters.Count - 1; i++)
            {
                if (System.Drawing.Printing.PrinterSettings.InstalledPrinters[i].ToString() == mPrinterName)
                {
                    return true;
                }
            }
            return false;
        }

        public static Image GetCode128(string sText)
        {
            // 定義產出是QR Code，Code128 就是 BarcodeFormat.CODE_128
            var writer = new BarcodeWriter();
            writer.Format = BarcodeFormat.CODE_128;

            // 定義長寛
            writer.Options.Height = 50;
            writer.Options.Width = 100;

            return (Image)writer.Write(sText);
        }
    }
}
