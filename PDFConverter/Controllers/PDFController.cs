

using Aspose.Pdf;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using ICSharpCode.SharpZipLib.Zip;

//using IronPdf;
//using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using Microsoft.Office.Interop.Word;


//using iText.IO.Image;
//using iText.Kernel.Pdf;
//using iText.Layout;
//using iText.Layout.Element;
using PDFConverter.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

using System.Web;
using System.Web.Mvc;
using static Aspose.Pdf.DocSaveOptions;
using PdfDocument = PdfSharp.Pdf.PdfDocument;
using SaveFormat = Aspose.Pdf.SaveFormat;

namespace PDFConverter.Controllers
{
    public class PDFController : Controller
    {
        public Microsoft.Office.Interop.Word.Document wordDocument { get; set; }


        public ActionResult PDFConverter()
        {
            ViewBag.Title = "PDF Converter - Convert to PDF Files Online FREE - PDFPuppy";

            return View();
        }
        public ActionResult FamilyProducts()
        {
            return View();
        }
        [HttpPost]
        public ActionResult ConvertWordtoPDF(IEnumerable<HttpPostedFileBase> product)
        {

            // Checking no of files injected in Request object  
            if (Request.Files.Count > 0)
            {
                try
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;
                    for (int i = 0; i < files.Count; i++)
                    {
                        //string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/";  
                        //string filename = Path.GetFileName(Request.Files[i].FileName);  

                        HttpPostedFileBase file = files[i];
                        string fname;

                        // Checking for Internet Explorer  
                        if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                        {
                            string[] testfiles = file.FileName.Split(new char[] { '\\' });
                            fname = testfiles[testfiles.Length - 1];
                        }
                        else
                        {
                            fname = file.FileName;
                        }

                        // Get the complete folder path and store the file inside it.  
                        var extension = System.IO.Path.GetExtension(file.FileName);
                        MemoryStream target = new MemoryStream();
                        file.InputStream.CopyTo(target);
                        byte[] data = target.ToArray();
                        var guidwordfilename = Guid.NewGuid().ToString() + extension;//".docx";
                        var guidpdffilename = Guid.NewGuid().ToString() + ".pdf";

                        // Get the complete folder path and store the file inside it.  
                        var wordfname = Path.Combine(Server.MapPath("~/Uploads/"), guidwordfilename);
                        var pdffname = Path.Combine(Server.MapPath("~/Uploads/"), guidpdffilename);
                        file.SaveAs(wordfname);
                        Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();

                        wordDocument = appWord.Documents.Open(wordfname);
                        wordDocument.ExportAsFixedFormat(pdffname, WdExportFormat.wdExportFormatPDF);
                        wordDocument.Close();

                        Marshal.ReleaseComObject(wordDocument);
                        Marshal.ReleaseComObject(appWord);

                        byte[] fileBytes = System.IO.File.ReadAllBytes(pdffname);
                        fileBytes = System.IO.File.ReadAllBytes(pdffname);
                        TempData["WordtoPDF"] = fileBytes;
                        FileInfo wordfile = new FileInfo(wordfname);
                        if (wordfile.Exists)
                        {
                            wordfile.Delete();
                        }

                        FileInfo pdffile = new FileInfo(pdffname);
                        if (pdffile.Exists)
                        {
                            pdffile.Delete();
                        }



                        // Returns message that successfully uploaded  

                    }
                    // Returns message that successfully uploaded  

                    return Json("Success");
                }
                catch (Exception ex)
                {
                    var st = new StackTrace(ex, true);
                    // Get the top stack frame
                    var frame = st.GetFrame(0);
                    // Get the line number from the stack frame
                    var line = frame.GetFileLineNumber();
                    return Json("Error occurred. Error details: " + ex.Message + "Line number :" + line);
                }
            }
            else
            {
                return Json("No files selected.");
            }
        }

        public ActionResult DownloadWordtoPDF()
        {
            // retrieve byte array here
            var array = TempData["WordtoPDF"] as byte[];
            if (array != null)
            {
                return File(array, System.Net.Mime.MediaTypeNames.Application.Octet, Guid.NewGuid().ToString() + ".pdf");
            }
            else
            {
                return new EmptyResult();
            }
        }
        public ActionResult DownloadImagetoPDF()
        {
            // retrieve byte array here
            var array = TempData["ImagetoPDF"] as byte[];
            if (array != null)
            {
                return File(array, System.Net.Mime.MediaTypeNames.Application.Octet, Guid.NewGuid().ToString() + ".pdf");
            }
            else
            {
                return new EmptyResult();
            }
        }
        [HttpPost]
        public ActionResult ConvertImageToPdf(IEnumerable<HttpPostedFileBase> product)
        {
            if (Request.Files.Count > 0)
            {
                try
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];
                        MemoryStream target = new MemoryStream();
                        file.InputStream.CopyTo(target);
                        byte[] data = target.ToArray();
                        var guidjpgfilename = Guid.NewGuid().ToString() + ".jpg";
                        var guidpdffilename = Guid.NewGuid().ToString() + ".pdf";
                        var guiddocxfilename = Guid.NewGuid().ToString() + ".docx";

                        // Get the complete folder path and store the file inside it.  
                        var jpgfname = Path.Combine(Server.MapPath("~/Uploads/"), guidjpgfilename);
                        var pdffname = Path.Combine(Server.MapPath("~/Uploads/"), guidpdffilename);
                        var docxfname = Path.Combine(Server.MapPath("~/Uploads/"), guiddocxfilename);
                        file.SaveAs(jpgfname);
                        //testing**********************
                        // Create Document and DocumentBuilder.
                        // The builder makes it simple to add content to the document.
                        Aspose.Words.Document doc = new Aspose.Words.Document();
                        DocumentBuilder builder = new DocumentBuilder(doc);


                        Aspose.Words.Drawing.Shape shape = builder.InsertImage(jpgfname);
                        shape.WrapType = WrapType.None;
                        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                        shape.HorizontalAlignment = Aspose.Words.Drawing.HorizontalAlignment.Center;
                        shape.VerticalAlignment = Aspose.Words.Drawing.VerticalAlignment.Center;

                        doc.Save(docxfname);

                        Aspose.Words.Document doc1 = new Aspose.Words.Document(docxfname);

                        foreach (Aspose.Words.Section section in doc1.Sections)
                        {
                            section.PageSetup.PaperSize = Aspose.Words.PaperSize.A4;
                            section.PageSetup.VerticalAlignment = PageVerticalAlignment.Center;


                        }

                        doc.UpdatePageLayout();
                        doc.Save(pdffname);
                        byte[] fileBytes = System.IO.File.ReadAllBytes(pdffname);
                        TempData["ImagetoPDF"] = fileBytes;

                        FileInfo jpgfile = new FileInfo(jpgfname);
                        if (jpgfile.Exists)
                        {
                            jpgfile.Delete();
                        }

                        FileInfo pdffile = new FileInfo(pdffname);
                        if (pdffile.Exists)
                        {
                            pdffile.Delete();
                        }
                        FileInfo docxfile = new FileInfo(docxfname);
                        if (docxfile.Exists)
                        {
                            docxfile.Delete();
                        }
                        // Returns message that successfully uploaded  
                        return Json("Success");
                    }
                }
                catch (Exception ex)
                {
                    return Json("Error occurred. Error details: " + ex.Message);
                }
            }
            return null;
        }





        public ActionResult PrivacyPolicy()
        {
            return View();
        }

        public ActionResult SupportProject()
        {
            return View();
        }

        public ActionResult DataProtectionSecurity()
        {
            return View();
        }

        public ActionResult Contact()
        {
            return View();
        }

        public ActionResult Feedback()
        {
            return View();
        }
        public ActionResult News()
        {
            return View();
        }




        [HttpPost]
        public ActionResult MergePDF(IEnumerable<HttpPostedFileBase> product)
        {
            var file1 = "";
            var fullfile1 = "";
            var file2 = "";
            var fullfile2 = "";
            var file3 = "";
            var fullfile3 = "";
            var file4 = "";
            var fullfile4 = "";
            var file5 = "";
            var fullfile5 = "";
            var mergedoutputfile = "";
            var fullmergedoutputfile = "";
            try
            {
                if (Request.Files.Count > 1)
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;
                    for (int i = 0; i < files.Count; i++)
                    {
                        HttpPostedFileBase file = files[i];
                        MemoryStream target = new MemoryStream();
                        file.InputStream.CopyTo(target);
                        byte[] data = target.ToArray();
                        if (i == 0)
                        {
                            file1 = Guid.NewGuid().ToString() + ".pdf";
                            fullfile1 = Path.Combine(Server.MapPath("~/Uploads/"), file1);
                            file.SaveAs(fullfile1);
                        }
                        else if (i == 1)
                        {
                            file2 = Guid.NewGuid().ToString() + ".pdf";
                            fullfile2 = Path.Combine(Server.MapPath("~/Uploads/"), file2);
                            file.SaveAs(fullfile2);
                        }
                        else if (i == 2)
                        {
                            file3 = Guid.NewGuid().ToString() + ".pdf";
                            fullfile3 = Path.Combine(Server.MapPath("~/Uploads/"), file3);
                            file.SaveAs(fullfile3);
                        }
                        else if (i == 3)
                        {
                            file4 = Guid.NewGuid().ToString() + ".pdf";
                            fullfile4 = Path.Combine(Server.MapPath("~/Uploads/"), file4);
                            file.SaveAs(fullfile4);
                        }
                        else if (i == 4)
                        {
                            file5 = Guid.NewGuid().ToString() + ".pdf";
                            fullfile5 = Path.Combine(Server.MapPath("~/Uploads/"), file5);
                            file.SaveAs(fullfile5);
                        }

                    }

                    using (PdfDocument one = file1 == string.Empty ? new PdfDocument() : PdfReader.Open(Server.MapPath("~/Uploads/" + file1), PdfDocumentOpenMode.Import))
                    using (PdfDocument two = file2 == string.Empty ? new PdfDocument() : PdfReader.Open(Server.MapPath("~/Uploads/" + file2), PdfDocumentOpenMode.Import))
                    using (PdfDocument three = file3 == string.Empty ? new PdfDocument() : PdfReader.Open(Server.MapPath("~/Uploads/" + file3), PdfDocumentOpenMode.Import))
                    using (PdfDocument four = file4 == string.Empty ? new PdfDocument() : PdfReader.Open(Server.MapPath("~/Uploads/" + file4), PdfDocumentOpenMode.Import))
                    using (PdfDocument five = file5 == string.Empty ? new PdfDocument() : PdfReader.Open(Server.MapPath("~/Uploads/" + file5), PdfDocumentOpenMode.Import))
                    using (PdfDocument outPdf = new PdfDocument())
                    {
                        CopyPages(one, outPdf);
                        CopyPages(two, outPdf);
                        CopyPages(three, outPdf);
                        CopyPages(four, outPdf);
                        CopyPages(five, outPdf);
                        mergedoutputfile = Guid.NewGuid().ToString() + ".pdf";
                        fullmergedoutputfile = Path.Combine(Server.MapPath("~/Uploads/"), mergedoutputfile);
                        outPdf.Save(Server.MapPath("~/Uploads/" + mergedoutputfile));
                    }
                    byte[] fileBytes = System.IO.File.ReadAllBytes(fullmergedoutputfile);
                    TempData["MergePDF"] = fileBytes;

                    FileInfo mergedfinalfile = new FileInfo(fullmergedoutputfile);
                    if (mergedfinalfile.Exists)
                    {
                        mergedfinalfile.Delete();
                    }
                    if (fullfile1 != string.Empty)
                    {
                        FileInfo file_1 = new FileInfo(fullfile1);
                        if (file_1.Exists)
                        {
                            file_1.Delete();
                        }
                    }
                    if (fullfile2 != string.Empty)
                    {
                        FileInfo file_2 = new FileInfo(fullfile2);
                        if (file_2.Exists)
                        {
                            file_2.Delete();
                        }
                    }
                    if (fullfile3 != string.Empty)
                    {
                        FileInfo file_3 = new FileInfo(fullfile3);
                        if (file_3.Exists)
                        {
                            file_3.Delete();
                        }
                    }
                    if (fullfile4 != string.Empty)
                    {
                        FileInfo file_4 = new FileInfo(fullfile4);
                        if (file_4.Exists)
                        {
                            file_4.Delete();
                        }
                    }
                    if (fullfile5 != string.Empty)
                    {
                        FileInfo file_5 = new FileInfo(fullfile5);
                        if (file_5.Exists)
                        {
                            file_5.Delete();
                        }
                    }
                    // Returns message that successfully uploaded  
                    return Json("Success");
                    // byte[] fileBytes = System.IO.File.ReadAllBytes(Server.MapPath("~/Uploads/" + mergedoutputfile));
                    //return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "~/Uploads/" + mergedoutputfile);
                }
            }
            catch (Exception ex)
            {
                return Json("Error occurred. Error details: " + ex.Message);
            }
            return null;
        }

        public ActionResult DownloadMergePDF()
        {
            // retrieve byte array here
            var array = TempData["MergePDF"] as byte[];
            if (array != null)
            {
                return File(array, System.Net.Mime.MediaTypeNames.Application.Octet, Guid.NewGuid().ToString() + ".pdf");
            }
            else
            {
                return new EmptyResult();
            }
        }

        void CopyPages(PdfDocument from, PdfDocument to)
        {
            if (from.PageCount > 0)
            {
                for (int i = 0; i < from.PageCount; i++)
                {
                    to.AddPage(from.Pages[i]);
                }
            }
        }

        [HttpPost]
        public ActionResult MergeTextToPDFBulk(IEnumerable<HttpPostedFileBase> product)
        {
            var pdffilename = "";
            var pdffilefullpath = "";
            var fullmergedoutputfile = "";
            var textfilefullpath = "";
            var guidtxtfilename = "";
            try
            {
                if (Request.Files.Count > 0)
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;
                    using (PdfDocument targetDoc = new PdfDocument())
                    {
                        for (int i = 0; i < files.Count; i++)
                        {
                            HttpPostedFileBase file = files[i];
                            MemoryStream target = new MemoryStream();
                            var extension = System.IO.Path.GetExtension(file.FileName);
                            if (extension.ToLower().Trim() == ".txt")
                            {
                                guidtxtfilename = Guid.NewGuid().ToString() + extension;//".docx";
                                file.InputStream.CopyTo(target);
                                byte[] data = target.ToArray();
                                textfilefullpath = Path.Combine(Server.MapPath("~/Uploads/"), guidtxtfilename);
                                pdffilename = Guid.NewGuid().ToString() + ".pdf";
                                pdffilefullpath = Path.Combine(Server.MapPath("~/Uploads/"), pdffilename);
                                file.SaveAs(textfilefullpath);
                                Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();

                                wordDocument = appWord.Documents.Open(textfilefullpath);
                                wordDocument.ExportAsFixedFormat(pdffilefullpath, WdExportFormat.wdExportFormatPDF);
                                wordDocument.Close();

                                Marshal.ReleaseComObject(wordDocument);
                                Marshal.ReleaseComObject(appWord);


                                using (PdfDocument pdfDoc = PdfReader.Open(pdffilefullpath, PdfDocumentOpenMode.Import))
                                {
                                    //for (int j = 0; j < pdfDoc.PageCount; j++)
                                    //{
                                    //    targetDoc.AddPage(pdfDoc.Pages[i]);
                                    //}

                                    foreach (var items in pdfDoc.Pages)
                                    {
                                        targetDoc.AddPage(items);
                                    }
                                }

                                FileInfo tempfile = new FileInfo(pdffilefullpath);
                                if (tempfile.Exists)
                                {
                                    tempfile.Delete();
                                }
                                FileInfo temptextfile = new FileInfo(textfilefullpath);
                                if (temptextfile.Exists)
                                {
                                    temptextfile.Delete();
                                }

                                fullmergedoutputfile = Path.Combine(Server.MapPath("~/Uploads/"), Guid.NewGuid().ToString() + ".pdf");
                                targetDoc.Save(fullmergedoutputfile);
                            }
                        }
                    }



                    byte[] fileBytes = System.IO.File.ReadAllBytes(fullmergedoutputfile);
                    TempData["MergeTextToPDFBulk"] = fileBytes;
                    FileInfo mergedfile = new FileInfo(fullmergedoutputfile);
                    if (mergedfile.Exists)
                    {
                        mergedfile.Delete();
                    }

                    // Returns message that successfully uploaded  
                    return Json("Success");
                    // byte[] fileBytes = System.IO.File.ReadAllBytes(Server.MapPath("~/Uploads/" + mergedoutputfile));
                    //return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "~/Uploads/" + mergedoutputfile);
                }
            }
            catch (Exception ex)
            {
                return Json("Error occurred. Error details: " + ex.Message);
            }
            return null;
        }

        public ActionResult DownloadMergetexttoPDFBulk()
        {
            // retrieve byte array here
            var array = TempData["MergeTextToPDFBulk"] as byte[];
            if (array != null)
            {
                return File(array, System.Net.Mime.MediaTypeNames.Application.Octet, Guid.NewGuid().ToString() + ".pdf");
            }
            else
            {
                return new EmptyResult();
            }
        }


        public ActionResult SplitPDF(IEnumerable<HttpPostedFileBase> product)

        {
            HttpFileCollectionBase files = Request.Files;
            HttpPostedFileBase file = files[0];
            MemoryStream target = new MemoryStream();
            var guidpdffilename = "";
            var guidpdftempfilename = "";
            var fileName = "";
            var tempOutPutPath = "";
            var extension = System.IO.Path.GetExtension(file.FileName);
            try
            {
                if (extension.ToLower().Trim() == ".pdf")
                {
                    // Input File name
                    guidpdffilename = Guid.NewGuid().ToString() + extension;//".docx";
                    file.InputStream.CopyTo(target);
                    byte[] data = target.ToArray();
                    guidpdftempfilename = Path.Combine(Server.MapPath("~/Uploads/"), guidpdffilename);
                    file.SaveAs(guidpdftempfilename);
                    // Input file path


                    // Open the input file in Import Mode
                    PdfDocument inputPDFFile = PdfReader.Open(guidpdftempfilename, PdfDocumentOpenMode.Import);

                    //Get the total pages in the PDF
                    var totalPagesInInputPDFFile = inputPDFFile.PageCount;
                    fileName = string.Format("{0}.zip", totalPagesInInputPDFFile.ToString() + "_Files_" + Guid.NewGuid().ToString());
                    tempOutPutPath = Server.MapPath("~/Uploads/Zip_PDF/") + fileName;

                    using (ZipOutputStream s = new ZipOutputStream(System.IO.File.Create(tempOutPutPath)))
                    {

                        s.SetLevel(9); // 0-9, 9 being the highest compression  

                        byte[] buffer = new byte[4096];
                        while (totalPagesInInputPDFFile != 0)
                        {

                            //Create an instance of the PDF document in memory
                            PdfDocument outputPDFDocument = new PdfDocument();

                            // Add a specific page to the PdfDocument instance
                            outputPDFDocument.AddPage(inputPDFFile.Pages[totalPagesInInputPDFFile - 1]);

                            //save the PDF document
                            string outputPDFFilePath = Path.Combine(Server.MapPath("~/Uploads/"), totalPagesInInputPDFFile.ToString() + ".pdf");
                            outputPDFDocument.Save(outputPDFFilePath);

                            ZipEntry entry = new ZipEntry(Path.GetFileName(outputPDFFilePath));
                            entry.DateTime = DateTime.Now;
                            entry.IsUnicodeText = true;
                            s.PutNextEntry(entry);
                            using (FileStream fs = System.IO.File.OpenRead(outputPDFFilePath))
                            {
                                int sourceBytes;
                                do
                                {
                                    sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                    s.Write(buffer, 0, sourceBytes);
                                } while (sourceBytes > 0);
                            }

                            totalPagesInInputPDFFile--;
                            FileInfo tempfile = new FileInfo(outputPDFFilePath);
                            if (tempfile.Exists)
                            {
                                tempfile.Delete();
                            }
                        }
                        s.Finish();
                        s.Flush();
                        s.Close();
                    }
                    byte[] finalResult = System.IO.File.ReadAllBytes(tempOutPutPath);
                    TempData["SplitPDFZip"] = finalResult;
                    TempData["ZipfileName"] = fileName;

                    FileInfo split_source_file = new FileInfo(guidpdftempfilename);
                    if (split_source_file.Exists)
                    {
                        split_source_file.Delete();
                    }
                    if (System.IO.File.Exists(tempOutPutPath))
                        System.IO.File.Delete(tempOutPutPath);


                    return Json("Success");
                }

            }
            catch (Exception ex)
            {
                return Json("Error occurred. Error details: " + ex.Message);
            }
            return null;

        }

        public ActionResult DownloadSplitPDFZip()
        {
            // retrieve byte array here
            var array = TempData["SplitPDFZip"] as byte[];
            var fileName = TempData["ZipfileName"] as string;
            if (array != null)
            {
                return File(array, "application/zip", fileName);
            }
            else
            {
                return new EmptyResult();
            }
        }


        public ActionResult PDFtoWord(IEnumerable<HttpPostedFileBase> product)
        {
            HttpFileCollectionBase files = Request.Files;
            HttpPostedFileBase file = files[0];
            MemoryStream target = new MemoryStream();
            var guidpdffilename = "";
            var guidpdftempfilename = "";
            var docfileName = "";
            var tempOutPutDocPath = "";
            var extension = System.IO.Path.GetExtension(file.FileName);
            try
            {
                if (extension.ToLower().Trim() == ".pdf")
                {
                    guidpdffilename = Guid.NewGuid().ToString() + extension;//".docx";
                    file.InputStream.CopyTo(target);
                    byte[] data = target.ToArray();
                    guidpdftempfilename = Path.Combine(Server.MapPath("~/Uploads/"), guidpdffilename);
                    file.SaveAs(guidpdftempfilename);
                    var document = new Aspose.Pdf.Document(guidpdftempfilename);
                    docfileName = Guid.NewGuid().ToString() + ".doc";
                    tempOutPutDocPath = Path.Combine(Server.MapPath("~/Uploads/"), docfileName);
                    // save document in DOC format

                    Aspose.Pdf.DocSaveOptions saveOptions = new Aspose.Pdf.DocSaveOptions
                    {
                        // Specify the output format as DOCX
                        Format =  DocFormat.DocX
                        // Set other DocSaveOptions params
                        // ....
                    };
                    //document.Save(tempOutPutDocPath, SaveFormat.Doc);
                    document.Save(tempOutPutDocPath, saveOptions);

                    Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();

                    wordDocument = appWord.Documents.Open(tempOutPutDocPath);

                    var count = wordDocument.Sections.Count;
                    for (var i = 0; i < count; i++)
                    {
                       // Microsoft.Office.Interop.Word.Words wds = wordDocument.Sections[1].Range.Words;


                        FindAndReplace(appWord, "Evaluation Only. Created with Aspose.PDF. Copyright 2002-2022 Aspose Pty Ltd.", "");

                    }
                    wordDocument.Save();
                   // wordDocument.SaveAs(tempOutPutDocPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                    wordDocument.Close();
                    appWord.Quit();
                    byte[] fileBytes = System.IO.File.ReadAllBytes(tempOutPutDocPath);
                    TempData["PDFtoWord"] = fileBytes;


                    FileInfo docfile = new FileInfo(tempOutPutDocPath);
                    if (docfile.Exists)
                    {
                        docfile.Delete();
                    }
                    FileInfo pdffile = new FileInfo(guidpdftempfilename);
                    if (pdffile.Exists)
                    {
                        pdffile.Delete();
                    }



                    return Json("Success");
                }
            }
            catch (Exception ex)
            {
                return Json("Error occurred. Error details: " + ex.Message);
            }
            return null;
        }


        public ActionResult DownloadPDFtoWord()
        {
            // retrieve byte array here
            var array = TempData["PDFtoWord"] as byte[];
            if (array != null)
            {
                return File(array, System.Net.Mime.MediaTypeNames.Application.Octet, Guid.NewGuid().ToString() + ".doc");
            }
            else
            {
                return new EmptyResult();
            }
        }
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application WordApp, object findText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object nmatchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
            object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
            WordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
            ref nmatchAllWordForms, ref forward,
            ref wrap, ref format, ref replaceWithText,
            ref replaceAll, ref matchKashida,
            ref matchDiacritics, ref matchAlefHamza,
            ref matchControl);
        }



        public ActionResult PDFtoWordDocx(IEnumerable<HttpPostedFileBase> product)
        {
            HttpFileCollectionBase files = Request.Files;
            HttpPostedFileBase file = files[0];
            MemoryStream target = new MemoryStream();
            var guidpdffilename = "";
            var guidpdftempfilename = "";
            var docfileName = "";
            var docxfileName = "";
            var tempOutPutDocPath = "";
            var tempOutPutDocxPath = "";
            var extension = System.IO.Path.GetExtension(file.FileName);
            try
            {
                if (extension.ToLower().Trim() == ".pdf")
                {
                    guidpdffilename = Guid.NewGuid().ToString() + extension;//".docx";
                    file.InputStream.CopyTo(target);
                    byte[] data = target.ToArray();
                    guidpdftempfilename = Path.Combine(Server.MapPath("~/Uploads/"), guidpdffilename);
                    file.SaveAs(guidpdftempfilename);
                    var document = new Aspose.Pdf.Document(guidpdftempfilename);
                    docfileName = Guid.NewGuid().ToString() + ".doc";
                    tempOutPutDocPath = Path.Combine(Server.MapPath("~/Uploads/"), docfileName);
                    docxfileName = Guid.NewGuid().ToString() + ".docx";
                    tempOutPutDocxPath = Path.Combine(Server.MapPath("~/Uploads/"), docxfileName);
                    // save document in DOC format
                    Aspose.Pdf.DocSaveOptions saveOptions = new Aspose.Pdf.DocSaveOptions
                    {
                        // Specify the output format as DOCX
                        Format = DocFormat.DocX
                        // Set other DocSaveOptions params
                        // ....
                    };
                    

                   document.Save(tempOutPutDocPath, saveOptions);
                     

                    Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();

                    wordDocument = appWord.Documents.Open(tempOutPutDocPath);

                    var count = wordDocument.Sections.Count;
                    for (var i = 0; i < count; i++)
                    {
                        // Microsoft.Office.Interop.Word.Words wds = wordDocument.Sections[1].Range.Words;


                        FindAndReplace(appWord, "Evaluation Only. Created with Aspose.PDF. Copyright 2002-2022 Aspose Pty Ltd.", "");

                    }
                    wordDocument.Save();
                    // wordDocument.SaveAs(tempOutPutDocPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                    wordDocument.Close();
                    wordDocument = appWord.Documents.Open(tempOutPutDocPath);
                    wordDocument.SaveAs(tempOutPutDocxPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                    wordDocument.Close();
                    appWord.Quit();



                    byte[] fileBytes = System.IO.File.ReadAllBytes(tempOutPutDocxPath);

                    TempData["PDFtoWorddocx"] = fileBytes;


                    FileInfo docfile = new FileInfo(tempOutPutDocPath);
                    if (docfile.Exists)
                    {
                        docfile.Delete();
                    }
                    FileInfo pdffile = new FileInfo(guidpdftempfilename);
                    if (pdffile.Exists)
                    {
                        pdffile.Delete();
                    }
                    FileInfo docxfile = new FileInfo(tempOutPutDocxPath);
                    if (docxfile.Exists)
                    {
                        docxfile.Delete();
                    }



                    return Json("Success");
                }
            }
            catch (Exception ex)
            {
                return Json("Error occurred. Error details: " + ex.Message);
            }
            return null;
        }


        public ActionResult DownloadPDFtoWordDocx()
        {
            // retrieve byte array here
            var array = TempData["PDFtoWorddocx"] as byte[];
            if (array != null)
            {
                return File(array, System.Net.Mime.MediaTypeNames.Application.Octet, Guid.NewGuid().ToString() + ".docx");
            }
            else
            {
                return new EmptyResult();
            }
        }
        public ActionResult CompressPDF()
        {


            return new EmptyResult();

        }
        public ActionResult DownloadCompressPDF()
        {
            // retrieve byte array here
            var array = TempData["CompressPDF"] as byte[];
            var fileName = TempData["CompressPDFName"] as string;
            if (array != null)
            {
                return File(array, "application/zip", fileName);
            }
            else
            {
                return new EmptyResult();
            }
        }



    }


    //private File ZipFile(int count)
    //{

    //    var fileName = string.Format("{0}_.zip",  count.ToString()+"_Files_"+ Guid.NewGuid().ToString());
    //    var tempOutPutPath = Server.MapPath(Url.Content("/TempPDF/")) + fileName;

    //    using (ZipOutputStream s = new ZipOutputStream(System.IO.File.Create(tempOutPutPath)))
    //    {
    //        s.SetLevel(9); // 0-9, 9 being the highest compression  

    //        byte[] buffer = new byte[4096];

    //        var ImageList = new List<string>();

    //        ImageList.Add(Server.MapPath("/Images/01.jpg"));
    //        ImageList.Add(Server.MapPath("/Images/02.jpg"));


    //        for (int i = 0; i < ImageList.Count; i++)
    //        {
    //            ZipEntry entry = new ZipEntry(Path.GetFileName(ImageList[i]));
    //            entry.DateTime = DateTime.Now;
    //            entry.IsUnicodeText = true;
    //            s.PutNextEntry(entry);

    //            using (FileStream fs = System.IO.File.OpenRead(ImageList[i]))
    //            {
    //                int sourceBytes;
    //                do
    //                {
    //                    sourceBytes = fs.Read(buffer, 0, buffer.Length);
    //                    s.Write(buffer, 0, sourceBytes);
    //                } while (sourceBytes > 0);
    //            }
    //        }
    //        s.Finish();
    //        s.Flush();
    //        s.Close();

    //    }

    //    byte[] finalResult = System.IO.File.ReadAllBytes(tempOutPutPath);
    //    if (System.IO.File.Exists(tempOutPutPath))
    //        System.IO.File.Delete(tempOutPutPath);

    //    if (finalResult == null || !finalResult.Any())
    //        throw new Exception(String.Format("No Files found with Image"));

    //    return File(finalResult, "application/zip", fileName);
    //}
}










