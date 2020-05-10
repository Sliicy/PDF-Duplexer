using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iText.Kernel.Pdf;

namespace PDF_Duplexer {
    public partial class FrmMain : Form {
        public FrmMain() {
            InitializeComponent();
        }

        public void SendToPrinter(string filePath, int pageCount) {
            Cursor.Current = Cursors.WaitCursor;
            ProcessStartInfo info = new ProcessStartInfo() {
                Verb = "print",
                FileName = filePath,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            Process p = new Process() {
                StartInfo = info
            };
            p.Start();
            p.WaitForInputIdle();

            // Wait minimum 3 seconds plus however many pages are being printed:
            if (pageCount < 3) {
                pageCount = 3;
            }
            System.Threading.Thread.Sleep(1000 * pageCount);

            if (false == p.CloseMainWindow())
                p.Kill();
            Cursor.Current = Cursors.Default;
        }

        private void CleanUp(string path1, string path2) {
            try {
                File.Delete(path1);
                File.Delete(path2);
            } catch {
                MessageBox.Show("Temporary PDF files couldn't be deleted!");
            }
        }
        private void Print() {
            // Protects accidentally deleting an existing file in Downloads:
            Random random = new Random();
            int randomNumber = random.Next(0, 10000);

            string destinationOdd = @"C:\Users\" + Environment.UserName + @"\Downloads\odd temp" + randomNumber.ToString() + ".pdf";
            string destinationEven = @"C:\Users\" + Environment.UserName + @"\Downloads\even temp" + randomNumber.ToString() + ".pdf";

            var dir = new DirectoryInfo(@"C:\Users\" + Environment.UserName + @"\Downloads");
            var pdfFile = (from f in dir.GetFiles()
                           orderby f.LastWriteTime
                           select f).Last();
            if (pdfFile.Extension.ToLower() != ".pdf") {
                MessageBox.Show("Please download a PDF first to print it double sided!");
                return;
            }
            DialogResult dialog1 = MessageBox.Show("Would you like to print " + pdfFile.Name + " double sided?", "Print " + pdfFile.Name + " double sided?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialog1 == DialogResult.Yes) {
                PdfDocument doc = new PdfDocument(new PdfReader(pdfFile.FullName));

                int numberOfPages = doc.GetNumberOfPages();

                List<int> odds = new List<int>();
                List<int> evens = new List<int>();
                for (var i = 1; i <= numberOfPages; i++) {
                    if (i % 2 != 0)
                        odds.Add(i);
                }
                for (int i = 1; i <= numberOfPages; i++) {
                    if (i % 2 == 0)
                        evens.Add(i);
                }
                evens.Reverse();

                PdfDocument oddPDFWriter = new PdfDocument(new PdfWriter(destinationOdd));
                PdfDocument evenPDFWriter = new PdfDocument(new PdfWriter(destinationEven));

                 //If odd number of pages, append blank page to beginning of even pages:
                if (numberOfPages % 2 != 0) {
                    evenPDFWriter.AddNewPage();
                }

                foreach (var oddPage in odds)
                    doc.CopyPagesTo(oddPage, oddPage, oddPDFWriter);
                foreach (var evenPage in evens)
                    doc.CopyPagesTo(evenPage, evenPage, evenPDFWriter);

                int oddPageCount = oddPDFWriter.GetNumberOfPages();
                int evenPageCount = evenPDFWriter.GetNumberOfPages();

                oddPDFWriter.Close();
                evenPDFWriter.Close();

                // Print odd side:
                SendToPrinter(destinationOdd, oddPageCount);

                DialogResult dialog2 = MessageBox.Show("Please put the papers back into the printer face down. Then press OK to continue.", "Put back into printer", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dialog2 != DialogResult.OK) {
                    CleanUp(destinationOdd, destinationEven);
                    return;
                }

                // Print even side:
                SendToPrinter(destinationEven, evenPageCount);

                // Clean up:
                CleanUp(destinationOdd, destinationEven);
                return;
            } else
                return;
        }

        private void BtnPrint_Click(object sender, EventArgs e) {
            Print();
        }

        private void BtnExit_Click(object sender, EventArgs e) {
            Application.Exit();
        }
    }
}
