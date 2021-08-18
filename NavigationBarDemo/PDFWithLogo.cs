using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDDL
{
    public class PDFWithLogo
    {

        public static string PreparedAndPrintedBy;

        public static void LogoEveryPage(string startFile, string watermarkedFile)
        {
            string appRootDir = new DirectoryInfo(Environment.CurrentDirectory).Parent.Parent.FullName;
            SaveFileDialog sfdd = new SaveFileDialog();


            PdfReader reader1 = new PdfReader(startFile);

            using (FileStream fs = new FileStream(watermarkedFile, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                PdfStamper stamper = new PdfStamper(reader1, fs);
                int pageCount = reader1.NumberOfPages;
                for (int i = 1; i <= pageCount; i++)
                {
                    iTextSharp.text.Rectangle rect = reader1.GetPageSize(i);
                    PdfContentByte cb = stamper.GetUnderContent(i);
                    iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(System.IO.Path.GetTempPath() + @"enviLOG Basic" + "\\enviro_logo.jpg");
                    jpg.ScaleToFit(100, 50);
                    jpg.SetAbsolutePosition(50, 800);
                    cb.AddImage(jpg, false);
                }
                stamper.Close();
            }
        }
    }
}
