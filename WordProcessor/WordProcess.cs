using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using System.Net.Http;
using System.Net.Http.Headers;
using Acrobat;
using System.Windows.Forms;
using System.Drawing;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Threading;
using _Application = Microsoft.Office.Interop.Word._Application;

namespace WordProcessorLib
{

    public class FileContent
    {
      
        public string FileName { get; set; }

        public byte[] bytes { get; set; }
        public ICollection<attachments> attachments { get; set; }

        public ICollection<Images> ImageList { get; set; }

    }
    public class attachments
    {
        public string FileName { get; set; }

        public byte[] bytes { get; set; }
    }

    public class Images
    {
        public string ImageName { get; set; }

        public byte[] Imagebytes { get; set; }
    }
    public class WordProcessor
    {
        //FileContent filecontnt = new FileContent();
        List<attachments> attachments = new List<attachments>();
        List<Images> ImageList = new List<Images>();
        _Application applicationclass = new Application();
       string attachmentsPath = HttpContext.Current.Server.MapPath("~/attachments/");

        public FileContent WordToHTML(HttpPostedFile file, object path)
        {
           
           
            Object missing = Missing.Value;
            var content = new MultipartContent();
            file.SaveAs(path.ToString());
            applicationclass.Documents.Open(ref path);
            applicationclass.Visible = false;
            object format = WdSaveFormat.wdFormatFilteredHTML;
            Document document = applicationclass.ActiveDocument;
            string Name = Path.GetFileName(file.FileName);

            object htmlFilePath = HttpContext.Current.Server.MapPath("~/") + Name + ".html";
            ExtractEmbdedObj(document);
            document.SaveAs2(ref htmlFilePath, ref format);

            //Close the word document.
            document.Close();
            applicationclass.Quit();
            FileContent filecontnt = new FileContent()
            {
                FileName=Name,
                bytes=File.ReadAllBytes(htmlFilePath.ToString()),
                attachments=attachmentcontent(),
               ImageList=ImageContent(Name)
                
            };

            //StreamContent filecontnt = CreateFileContent(htmlFilePath.ToString(), file.FileName, ".html");
            // content.Add(filecontnt);
            return filecontnt;
        
            }
      
        public void ExtractEmbdedObj(Document document)
        {
            object VerbIndex = 1;
            Object missing = System.Reflection.Missing.Value;
            if (document.InlineShapes.Count > 0)
            {


                foreach (InlineShape inlineShape in document.InlineShapes)

                {
                    if(inlineShape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        Thread thread = new Thread(CopyFromClipbordInlineShape);
                        thread.SetApartmentState(ApartmentState.STA);
                        thread.Start();
                        thread.Join();
                        inlineShape.Delete();
                       
                    }
                   else if (inlineShape.OLEFormat != null && inlineShape.OLEFormat.ProgID != null)

                    {
                        
                        
                        //string attachmentsPath = HttpContext.Current.Server.MapPath("~/attachments/");
                        Directory.CreateDirectory(attachmentsPath);

                        switch (inlineShape.OLEFormat.ProgID)

                        {

                            case "PowerPoint.Show.8":

                                inlineShape.OLEFormat.DoVerb(ref VerbIndex);

                                Microsoft.Office.Interop.PowerPoint.Application ppt = Marshal.GetActiveObject("PowerPoint.Application") as Microsoft.Office.Interop.PowerPoint.Application;

                                ppt.ActivePresentation.SaveAs(attachmentsPath + "testPPT.pptx", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPresentation, Microsoft.Office.Core.MsoTriState.msoTrue);
                               
                                ppt.Quit();

                               

                             
                                inlineShape.Delete();
                                break;

                            case "Excel.Sheet.8":

                                inlineShape.OLEFormat.DoVerb(ref VerbIndex);

                                Microsoft.Office.Interop.Excel.Application excel = Marshal.GetActiveObject("Excel.Application") as Microsoft.Office.Interop.Excel.Application;

                                excel.ActiveWorkbook.SaveAs(attachmentsPath + inlineShape.OLEFormat.IconLabel, missing, missing, missing, missing, missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missing,

                                    missing, missing, missing, missing);
                                

                                excel.Workbooks.Close();

                                excel.Quit();
                               
                               
                                inlineShape.Delete();
                             
                                break;

                            case "Word.Document.8":

                                Microsoft.Office.Interop.Word.Document document1 = inlineShape.OLEFormat.Object as Microsoft.Office.Interop.Word.Document;

                                if (inlineShape.OLEFormat.IconLabel != null && inlineShape.OLEFormat.IconLabel != "")
                                {
                                    object fileName = attachmentsPath + inlineShape.OLEFormat.IconLabel;

                                    document1.SaveAs(ref fileName, ref missing, ref missing, ref missing, ref missing, ref missing,

                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,

                                        ref missing, ref missing, ref missing);
                                }
                               
                              
                                inlineShape.Delete();
                                break;

                            case "Package":
                                object verb = Microsoft.Office.Interop.Word.WdOLEVerb.wdOLEVerbShow;

                                inlineShape.OLEFormat.DoVerb(ref verb);
                                Microsoft.Office.Interop.Outlook.Application app = Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                               
                                Object selObject = app.ActiveInspector().CurrentItem;
                                Microsoft.Office.Interop.Outlook.MailItem mailItem = (selObject as Microsoft.Office.Interop.Outlook.MailItem);
                                String filepath = attachmentsPath + inlineShape.OLEFormat.IconLabel;
                                mailItem.SaveAs(filepath, OlSaveAsType.olMSG);
                                break;
                            /*case "Package":
                            inlineShape.OLEFormat.Activate();
                            //var pdfobject = inlineShape.OLEFormat.Object;

                            string filePath = attachmentsPath + inlineShape.OLEFormat.IconLabel;
                            var pExportFormat = WdExportFormat.wdExportFormatPDF;
                            bool pOpenAfterExport = false;
                            var pExportOptimizeFor = WdExportOptimizeFor.wdExportOptimizeForOnScreen;
                            var pExportRange = Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument;
                            int pStartPage = 0;
                            int pEndPage = 0;
                            var pExportItem = Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent;
                            var pIncludeDocProps = true;
                            var pKeepIRM = true;
                            var pCreateBookmarks = Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                            var pDocStructureTags = true;
                            var pBitmapMissingFonts = true;
                            var pUseISO19005_1 = false;
                            //var pdf = inlineShape.OLEFormat.Object as Acrobat.;

                            //inlineShape.OLEFormat.ConvertTo(inlineShape.OLEFormat.ClassType);
                            inlineShape.OLEFormat.Object.SaveAs2(filePath, WdExportFormat.wdExportFormatPDF);

                            //document2.ExportAsFixedFormat(filePath, pExportFormat, pOpenAfterExport, pExportOptimizeFor, pExportRange, pStartPage, pEndPage, pExportItem, pIncludeDocProps, pKeepIRM, pCreateBookmarks, pDocStructureTags, pBitmapMissingFonts, pUseISO19005_1);


                            inlineShape.Delete();
                            break;*/

                            default:

                                break;

                        }

                    }

                  

                }
            }
        }
        protected void CopyFromClipbordInlineShape()
        {
            foreach (InlineShape inlineShape in applicationclass.ActiveDocument.InlineShapes)

            {
                if (inlineShape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    // InlineShape inlineShape = applicationclass.ActiveDocument.InlineShapes[m_i];
                    inlineShape.Select();
                    applicationclass.Selection.Copy();
                    //Computer computer = new Computer();
                    //Image img = computer.Clipboard.GetImage();
                    if (Clipboard.GetDataObject() != null)
                    {
                        System.Windows.Forms.IDataObject data = Clipboard.GetDataObject();
                        if (data.GetDataPresent(System.Windows.Forms.DataFormats.Bitmap))
                        {
                            Image image = (Image)data.GetData(System.Windows.Forms.DataFormats.Bitmap, true);
                            image.Save(attachmentsPath+"image.gif", System.Drawing.Imaging.ImageFormat.Gif);
                            //image.Save(HttpContext.Current.Server.MapPath("~/attacments/image.jpg"), System.Drawing.Imaging.ImageFormat.Jpeg);

                        }
                        else
                        {
                            //LabelMessage.Text = "The Data In Clipboard is not as image format";
                        }
                    }
                    else
                    {
                        //LabelMessage.Text = "The Clipboard was empty";
                    }
                }
            }
        }
        public StreamContent CreateFileContent(string path, string fileName, string extention)
        {
            var stream = new FileStream(path, FileMode.Open);
            var fileContent = new StreamContent(stream);
            fileContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                Name = "\"files\"",
                FileName = "\"" + fileName + "\""
            };
            string contentType = GetContentType(extention);
            fileContent.Headers.ContentType = new MediaTypeHeaderValue(contentType);
            return fileContent;
        }

        public List<attachments> attachmentcontent()
        {
            var searchDirectory = new System.IO.DirectoryInfo(HttpContext.Current.Server.MapPath("~/attachments/"));

            foreach (var file in searchDirectory.GetFiles())
            {
                attachments.Add(new attachments
                {
                    FileName = file.FullName,
                    bytes = File.ReadAllBytes(file.FullName)
                });
            }
            return attachments;
        }

        public List<Images> ImageContent(string Name)
        {
            var ImageDirectory = new System.IO.DirectoryInfo(HttpContext.Current.Server.MapPath("~/") + Name + "_files");

            foreach (var img in ImageDirectory.GetFiles())
            {
                ImageList.Add(new Images
                {
                    ImageName = img.FullName,
                    Imagebytes = File.ReadAllBytes(img.FullName)
                });
            }
            return ImageList;
        }
        public string GetContentType(string extention)
        {
            switch (extention)
            {

                case ".xls":
                    return "application/vnd.ms-excel";

                case ".pptx":
                    return "application/vnd.openxmlformats-officedocument.presentationml.presentation";

                case ".doc":
                    return "application/msword";

                case ".html":
                    return "text/html";

                default:
                    return "";

            }

        }

    }


}
