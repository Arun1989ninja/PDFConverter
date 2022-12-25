using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Microsoft.Office.Interop.Word;
//using iText.IO.Image;
//using iText.Kernel.Pdf;
//using iText.Layout;
//using iText.Layout.Element;
using PDFConverter.Models;
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


namespace PDFConverter.Controllers
{
    public class HomeController : Controller
    {
        public Microsoft.Office.Interop.Word.Document wordDocument { get; set; }
        public ActionResult Index()
        {
            ViewBag.Title = "PDF PUPPY";

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
                    return Json("Error occurred. Error details: " + ex.Message +"Line number :"+ line);
                }
            }
            else
            {
                return Json("No files selected.");
            }
        }
        //[HttpPost]
        //public ActionResult ConvertImagetoPDF(IEnumerable<HttpPostedFileBase> product)
        //{



        //    // Checking no of files injected in Request object  
        //    if (Request.Files.Count > 0)
        //    {
        //        try
        //        {
        //            //  Get all files from Request object  
        //            HttpFileCollectionBase files = Request.Files;
        //            for (int i = 0; i < files.Count; i++)
        //            {

        //                HttpPostedFileBase file = files[i];
        //                MemoryStream target = new MemoryStream();
        //                file.InputStream.CopyTo(target);
        //                byte[] data = target.ToArray();
        //                var guidjpgfilename = Guid.NewGuid().ToString() + ".jpg";
        //                var guidpdffilename = Guid.NewGuid().ToString() + ".pdf";

        //                // Get the complete folder path and store the file inside it.  
        //                var jpgfname = Path.Combine(Server.MapPath("~/Uploads/"), guidjpgfilename);
        //                var pdffname = Path.Combine(Server.MapPath("~/Uploads/"), guidjpgfilename);
        //                file.SaveAs(jpgfname);
        //                ImageData imageData = ImageDataFactory.Create(jpgfname);
        //                PdfDocument pdfDocument = new PdfDocument(new PdfWriter(pdffname));

        //                Document document = new Document(pdfDocument);

        //                Image image = new Image(imageData);
        //                image.SetWidth(pdfDocument.GetDefaultPageSize().GetWidth() - 65);
        //                //image.SetAutoScaleHeight(true);
        //                image.SetHeight(pdfDocument.GetDefaultPageSize().GetHeight() - 65);

        //                document.Add(image);
        //                pdfDocument.Close();

        //                byte[] fileBytes = System.IO.File.ReadAllBytes(pdffname);
        //                TempData["ImagetoPDF"] = fileBytes;



        //                // Returns message that successfully uploaded  
        //                return Json("Success");
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            return Json("Error occurred. Error details: " + ex.Message);
        //        }

        //    }
        //    return null;

        //}
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
                        shape.HorizontalAlignment = HorizontalAlignment.Center;
                        shape.VerticalAlignment = VerticalAlignment.Center;                        

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





        [HttpPost]
        public ActionResult ConvertTiffAndGifToPdf(IEnumerable<HttpPostedFileBase> product)
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
                        file.SaveAs(jpgfname);
                        //testing**********************
                        // Create Document and DocumentBuilder.
                        // The builder makes it simple to add content to the document.
                        Aspose.Words.Document doc = new Aspose.Words.Document();
                        DocumentBuilder builder = new DocumentBuilder(doc);

                        //Read the image from file, ensure it is disposed.
                        using (Image image = Image.FromFile(jpgfname))
                        {
                            // Find which dimension the frames in this image represent. For example
                            // The frames of a BMP or TIFF are "page dimension" whereas frames of a GIF image are "time dimension". 
                            FrameDimension dimension = new FrameDimension(image.FrameDimensionsList[0]);

                            // Get the number of frames in the image.
                            int framesCount = image.GetFrameCount(dimension);

                            // Loop through all frames.
                            for (int frameIdx = 0; frameIdx < framesCount; frameIdx++)
                            {
                                // Insert a section break before each new page, in case of a multi-frame TIFF.
                                if (frameIdx != 0)
                                    builder.InsertBreak(BreakType.SectionBreakNewPage);

                                // Select active frame.
                                image.SelectActiveFrame(dimension, frameIdx);




                                // We want the size of the page to be the same as the size of the image.
                                // Convert pixels to points to size the page to the actual image size.
                                Aspose.Words.PageSetup ps = builder.PageSetup;
                                ps.PageWidth = ConvertUtil.PixelToPoint(image.Width, image.HorizontalResolution);
                                ps.PageHeight = ConvertUtil.PixelToPoint(image.Height, image.VerticalResolution);



                                // Insert the image into the document and position it at the top left corner of the page.
                                var result = builder.InsertImage(
                                       image,
                                       RelativeHorizontalPosition.Page,
                                       0,
                                       RelativeVerticalPosition.Page,
                                       0,
                                       ps.PageWidth,
                                           ps.PageHeight,
                                           WrapType.None);
                                result.HorizontalAlignment = HorizontalAlignment.Center;
                                result.VerticalAlignment = VerticalAlignment.Center;






                            }



                        }

                        //Save the document to PDF.
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

    }


}
