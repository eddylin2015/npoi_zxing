using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using ZXing.Common;
using ZXing;
using ZXing.Client.Result;
using ZXing.QrCode;
using System.Collections.Generic;

namespace npoi_ex1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("Hello World!");
            Console.WriteLine("GET api/NpoiXls/Files");
            ArrayList sb = new System.Collections.ArrayList();
            String path = @"c:\code\npoi_ex\159.xlsx";
            String filename = "159.xlsx";

            esNpoi.readxls(path, filename, "0", sb);
            Console.WriteLine(sb.Count);
            foreach (Object obj in sb)
            {
                /*
                Console.Write("(");
                foreach (Object sobj in (ArrayList)obj)

                    Console.Write("   '{0}',", sobj.ToString().Replace('\n',' ').Replace('\'',' '));
                Console.WriteLine("),");
                */
                ArrayList aobj = (ArrayList)obj;
                if (aobj.Count > 6 && aobj[6].ToString().Length > 6)
                {
                    string outpath = @"c:\code\npoi_ex\out\" + aobj[6].ToString() + ".jpg";
                    Console.WriteLine(aobj[6].ToString());
                    //barcode128(aobj[6].ToString());
                    //barcode(aobj[6].ToString(),80,420,10,BarcodeFormat.CODE_128,outpath);
                   // barcode(aobj[6].ToString(),250,250,0,BarcodeFormat.QR_CODE,outpath);
                    //barcode(aobj[6].ToString(), 250, 250, 0, BarcodeFormat.DATA_MATRIX,outpath);
                    barcode_decode(outpath);
                }
            }






            //return sb;
        }
        static void barcode_decode(string path)
        {
            ImageConverter converter = new ImageConverter();

            var reader = new ZXing.QrCode.QRCodeReader();
            //ZXing.BinaryBitmap 
            var dest = (Bitmap)Bitmap.FromFile(path);
            var bb=(byte[])converter.ConvertTo(dest, typeof(byte[]));
            // ;
            // BinaryBitmap bitmap = new BinaryBitmap(new HybridBinarizer(source));
            var source = new BitmapLuminanceSource(dest);
           // LuminanceSource source = new RGBLuminanceSource(bb,dest.Width,dest.Height, RGBLuminanceSource.BitmapFormat.RGB32);
            HybridBinarizer binarizer = new HybridBinarizer(source);
            BinaryBitmap binBitmap = new BinaryBitmap(binarizer);
            QRCodeReader qrr = new QRCodeReader();
            var result = qrr.decode(binBitmap);
            if (result != null)
            {

                var text = result.Text;
                Console.WriteLine(text);
            }
            else
            {
                Console.WriteLine("Err");
            }

        }
        static void barcode(String x, int h, int w, int m, BarcodeFormat bf, string outpath)
        {
            BarcodeWriterPixelData writer = new BarcodeWriterPixelData()
            {
                Format = bf,//BarcodeFormat.CODE_128,
                Options = new EncodingOptions
                {
                    Height = h,
                    Width = w,
                    PureBarcode = false, // this should indicate that the text should be displayed, in theory. Makes no difference, though.
                    Margin = m
                }
            };
            var pixelData = writer.Write(x);

            using (var bitmap = new Bitmap(pixelData.Width, pixelData.Height, System.Drawing.Imaging.PixelFormat.Format32bppRgb))
            {
                var bitmapData = bitmap.LockBits(new Rectangle(0, 0, pixelData.Width, pixelData.Height), System.Drawing.Imaging.ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppRgb);
                try
                {
                    System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bitmapData.Scan0, pixelData.Pixels.Length);
                }
                finally
                {
                    bitmap.UnlockBits(bitmapData);
                }
                bitmap.Save(outpath, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
        }

        static void barcode128(String x)
        {
            BarcodeWriterPixelData writer = new BarcodeWriterPixelData()
            {
                Format = BarcodeFormat.CODE_128,
                Options = new EncodingOptions
                {
                    Height = 80,
                    Width = 420,
                    PureBarcode = false, // this should indicate that the text should be displayed, in theory. Makes no difference, though.
                    Margin = 10
                }
            };
            var pixelData = writer.Write(x);

            using (var bitmap = new Bitmap(pixelData.Width, pixelData.Height, System.Drawing.Imaging.PixelFormat.Format32bppRgb))
            {
                var bitmapData = bitmap.LockBits(new Rectangle(0, 0, pixelData.Width, pixelData.Height), System.Drawing.Imaging.ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppRgb);
                try
                {
                    System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bitmapData.Scan0, pixelData.Pixels.Length);
                }
                finally
                {
                    bitmap.UnlockBits(bitmapData);
                }

                bitmap.Save(@"c:\code\npoi_ex\out\" + x + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                /*
				using (var ms = new System.IO.MemoryStream())
                {
                    var bitmapData = bitmap.LockBits(new Rectangle(0, 0, pixelData.Width, pixelData.Height), System.Drawing.Imaging.ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppRgb);
                    try
                    {
                        System.Runtime.InteropServices.Marshal.Copy(pixelData.Pixels, 0, bitmapData.Scan0, pixelData.Pixels.Length);
                    }
                    finally
                    {
                        bitmap.UnlockBits(bitmapData);
                    }       
					
                    return File(ms.ToArray(), "image/jpeg");
                }*/
            }
        }
    }
    public class esNpoi
    {
        public static void CellVal(ICell c, String v)
        {
            Regex NumberPattern = new Regex(@"^(-?\d+)(\.\d+)?$");
            if (NumberPattern.IsMatch(v))
            {
                c.SetCellValue(Double.Parse(v));
            }
            else
                c.SetCellValue(v);
        }

        public static void readxls(String filepath, String originalname, String sheetName, ArrayList sb)
        {
            // Note:  filepath = @"C:\code\sharefiles\" +userid+ @"\" + filename ;            
            if (!(filepath.ToUpper().EndsWith(".XLS") || filepath.ToUpper().EndsWith(".XLSX") || filepath.ToUpper().EndsWith(".XLSM"))) return;
            if (!File.Exists(filepath)) return;
            ISheet sheet = null;
            if (filepath.ToUpper().EndsWith(".XLS"))
            {
                HSSFWorkbook hssfwb;
                using (FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Read)) { hssfwb = new HSSFWorkbook(file); }
                string pattern = @"^\d+$";
                Match m = Regex.Match(sheetName, pattern, RegexOptions.IgnoreCase);
                if (m.Success)
                {
                    int sheetindex = int.Parse(sheetName);
                    sheet = hssfwb.GetSheetAt(sheetindex);
                }
                else
                {
                    sheet = hssfwb.GetSheet(sheetName);
                }
            }
            else
            {
                XSSFWorkbook hssfwb;
                using (FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Read)) { hssfwb = new XSSFWorkbook(file); }
                string pattern = @"^\d+$";
                Match m = Regex.Match(sheetName, pattern, RegexOptions.IgnoreCase);
                if (m.Success)
                {
                    int sheetindex = int.Parse(sheetName);
                    sheet = hssfwb.GetSheetAt(sheetindex);
                }
                else
                {
                    sheet = hssfwb.GetSheet(sheetName);
                }
            }
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null)
                {
                    ArrayList items = new ArrayList();
                    //for (int col = 0; col < sheet.GetRow(row).Cells.Count; col++)
                    for (int col = 0; col < sheet.GetRow(row).LastCellNum; col++)
                    {
                        var cell = sheet.GetRow(row).GetCell(col);
                        if (cell == null)
                        {
                            items.Add("");
                            continue;
                        }
                        var val = "";
                        switch (cell.CellType)
                        {
                            case CellType.Numeric:
                                if (HSSFDateUtil.IsCellDateFormatted(cell))
                                {
                                    val = (cell.DateCellValue.Date.Year == 1899) ? cell.DateCellValue.ToString("HH:mm", CultureInfo.InvariantCulture) : cell.DateCellValue.Date.ToString("yyyy-MM-dd");
                                }
                                else
                                    val = cell.NumericCellValue.ToString();
                                break;
                            case CellType.String:
                                val = cell.StringCellValue;
                                break;
                            case CellType.Blank:
                                val = string.Empty;
                                break;
                            case CellType.Formula:
                                val = cell.NumericCellValue.ToString();
                                break;
                        }
                        items.Add(val);
                    }
                    sb.Add(items);
                }
            }
            // Save the file
            // using (FileStream file = new FileStream(filepath, FileMode.Open, FileAccess.Write)) { hssfwb.Write(file); }            
            return;
        }
    }
}
