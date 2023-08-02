using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using PuppeteerSharp;
using PuppeteerSharp.Media;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace HTMLtoWord
{
    public partial class SamplePage : System.Web.UI.Page
    {
        
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        private static void SaveDOCX(string fileName, string BodyText, bool isLandScape, double rMargin, double lMargin, double bMargin, double tMargin)
        {
            
            WordprocessingDocument document = WordprocessingDocument.Open("E:\\HTML to Word converter\\sample.docx", true);
            MainDocumentPart mainDocumenPart = document.MainDocumentPart;

            //Place the HTML String into a MemoryStream Object
            using (StreamReader Reader = new StreamReader("E:\\HTML to Word converter\\Audit Report Tempalte (2).html"))
            {
                StringBuilder Sb = new StringBuilder();
                Sb.Append(Reader.ReadToEnd());
                MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(Sb.ToString()));

                //Assign an HTML Section for the String Text
                string htmlSectionID = "Sect1";

                // Create alternative format import part.
                AlternativeFormatImportPart formatImportPart = mainDocumenPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, htmlSectionID);

                // Feed HTML data into format import part (chunk).
                formatImportPart.FeedData(ms);
                AltChunk altChunk = new AltChunk();
                altChunk.Id = htmlSectionID;

                //Clear out the Document Body and Insert just the HTML string.  (This prevents an empty First Line)
                mainDocumenPart.Document.Body.RemoveAllChildren();
                mainDocumenPart.Document.Body.Append(altChunk);

                /*
                 Set the Page Orientation and Margins Based on Page Size
                 inch equiv = 1440 (1 inch margin)
                 */
                double width = 8.5 * 1440;
                double height = 11 * 1440;

                SectionProperties sectionProps = new SectionProperties();
                PageSize pageSize;
                if (isLandScape)
                    pageSize = new PageSize() { Width = (UInt32Value)height, Height = (UInt32Value)width, Orient = PageOrientationValues.Landscape };
                else
                    pageSize = new PageSize() { Width = (UInt32Value)width, Height = (UInt32Value)height, Orient = PageOrientationValues.Portrait };

                rMargin = rMargin * 1440;
                lMargin = lMargin * 1440;
                bMargin = bMargin * 1440;
                tMargin = tMargin * 1440;

                PageMargin pageMargin = new PageMargin() { Top = (Int32)tMargin, Right = (UInt32Value)rMargin, Bottom = (Int32)bMargin, Left = (UInt32Value)lMargin, Header = (UInt32Value)360U, Footer = (UInt32Value)360U, Gutter = (UInt32Value)0U };

                sectionProps.Append(pageSize);
                sectionProps.Append(pageMargin);
                mainDocumenPart.Document.Body.Append(sectionProps);

                //Saving/Disposing of the created word Document
                document.MainDocumentPart.Document.Save();
                document.Dispose();
            }
        }
        protected async void Button1_Click(object sender, EventArgs e)
        {
           // SaveDOCX("abc", null, false, 0, 0, 0,0);
            string _fileCSS = Server.MapPath("~/css/style.css");
            string _strCSS = File.ReadAllText(_fileCSS);
            string _baseURL = "http://localhost:1385/";
            string _filename = System.Guid.NewGuid().ToString() + ".doc";
            string htmlRaw = @"<table class='tbl'><thead><tr><th class='style0' colspan='2'> <img src='" + _baseURL + "img/logo.png' style='width: 180px;' /></th><th class='style1' colspan='4'><p style='font-size: 24px; padding-bottom: 2px; padding-top: 2px; font-weight: bold; margin-bottom: 1px;'>INVOICE</p> ID-2021-0024<br> Issue Date:21/09/2021<br> Delivery Date: 22/09/2021<br> Due Date:30/09/2021<br> <br><p style='font-size: 24px; padding-bottom: 2px; padding-top: 2px; font-weight: bold; margin-bottom: 1px;'>CLIENT DETAILS</p> Client 1<br> GST Number:XXXXXXXXXX</th></tr></thead><tbody><tr><td class='headstyle0' colspan='5' style='padding-top: 60px;'></td></tr><tr><td class='style3a'>ITEM</td><td class='style3a'>DESCRIPTION</td><td class='style3a'>QUANTITY</td><td class='style3a'>UNIT PRICE</td><td class='style3a'>TOTAL</td></tr><tr><td class='style3'>Item-1</td><td class='style3'>Description -1</td><td class='style3'>2 Pkt</td><td class='style3'>90.00</td><td class='style3b'>180.00</td></tr><tr><td class='style3'>Item-2</td><td class='style3'>Description-2</td><td class='style3'>5 Pkt</td><td class='style3'>35.00</td><td class='style3b'>175.00</td></tr><tr><td class='style3'>Item-3</td><td class='style3'>Description-3</td><td class='style3'>5 Kg</td><td class='style3'>50.00</td><td class='style3b'>250.00</td></tr><tr><td class='style3'>Item-4</td><td class='style3'>Description-4</td><td class='style3'>5 Kg</td><td class='style3'>150.00</td><td class='style3b'>750.00</td></tr><tr><td class='style3'>Item-5</td><td class='style3'>Description-5</td><td class='style3'>5 Kg</td><td class='style3'>100.00</td><td class='style3b'>500.00</td></tr><tr><td class='style0' colspan='2' rowspan='3'></td><td class='style3' colspan='2'>Total</td><td class='style3b'>1855.00</td></tr><tr><td class='style3' colspan='2'>GST@18%</td><td class='style3b'>333.90</td></tr><tr><td class='style3' colspan='2'>Net Payable Amount</td><td class='style3b'>2188.90</td></tr><tr><td class='style0' colspan='5' style='padding-top: 100px;'></td></tr><tr><td class='style0' colspan='5' style='background-color: aliceblue; border-radius: 2px;'><i>Note:Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</i></td></tr><tr><td class='style1' colspan='5' style='padding-top: 150px;'> Thank You<br> <b>CodeSample</b></td></tr></tbody></table>";

            StringBuilder strHTML = new StringBuilder("");
            strHTML.Append("<html " +
                " xmlns:o='urn:schemas-microsoft-com:office:office'" +
                " xmlns:w='urn:schemas-microsoft-com:office:word'" +
                " xmlns='http://www.w3.org/TR/REC-html40'>" +
                "<head><title>Invoice Sample</title>");

            strHTML.Append("<xml><w:WordDocument>" +
                " <w:View>Print</w:View>" +
                " <w:Zoom>100</w:Zoom>" +
                " <w:DoNotOptimizeForBrowser/>" +
                " </w:WordDocument>" +
                " </xml>");

            strHTML.Append("<style>" + _strCSS + "</style></head>");
            strHTML.Append("<body><div class='page-settings'>" + htmlRaw + "</div></body></html>");
            var file = Server.MapPath("XYZ Corporation_template.html");
            await ReplaceSectionWithImage1(file, "title");
            using (StreamReader Reader = new StreamReader(file))
            {
                StringBuilder Sb = new StringBuilder();
                Sb.Append(Reader.ReadToEnd());
                
           
            Response.AppendHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
            Response.AppendHeader("Content-disposition", "attachment;filename=" + _filename + "");
            Response.Write(Sb.ToString());
            }
        }
        public static string ExtractSectionById(string htmlFilePath, string sectionId)
        {
            string htmlContent = File.ReadAllText(htmlFilePath);
            var htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(htmlContent);

            var sectionNode = htmlDoc.GetElementbyId(sectionId);
            if (sectionNode != null)
            {
                return sectionNode.OuterHtml;
            }

            return null;
        }
        public static async Task<string> HtmlToImage(string htmlContent, int width, int height)
        {
            var browserFetcher = new BrowserFetcher();
            await browserFetcher.DownloadAsync(BrowserFetcher.DefaultChromiumRevision);

            var browser = await Puppeteer.LaunchAsync(new LaunchOptions
            {
               
            });

            var page = await browser.NewPageAsync();
            await page.SetContentAsync(htmlContent);
            await page.SetViewportAsync(new ViewPortOptions
            {
                Width = width,
                Height = height
            });
            var sectionElementHandle = await page.QuerySelectorAsync("#title");
           // if (sectionElementHandle != null)
           // {
                // Capture the screenshot of the section only
                var image = await sectionElementHandle.ScreenshotDataAsync(new ScreenshotOptions
                {
                    FullPage = false
                });
           // }
                var screenshot = await page.ScreenshotDataAsync();
            await browser.CloseAsync();

            return Convert.ToBase64String(image);
        }
        public static async Task ReplaceSectionWithImage(string htmlFilePath, string sectionId)
        {
            string extractedSection = ExtractSectionById(htmlFilePath, sectionId);
            if (extractedSection != null)
            {
                // Convert the extracted section to an image
                int imageWidth = 40;
                int imageHeight = 40;
                string base64Image = await HtmlToImage(extractedSection, imageWidth, imageHeight);

                // Replace the section with an image tag in the original HTML
                string newHtmlContent = File.ReadAllText(htmlFilePath);
                newHtmlContent = newHtmlContent.Replace(extractedSection, $"<img src=\"data:image/png;base64,{base64Image}\" alt=\"{sectionId}\">");
                File.WriteAllText(htmlFilePath, newHtmlContent);
            }
        }


public static async Task ReplaceSectionWithImage1(string htmlFilePath, string sectionId)
    {
            var a = System.IO.Path.GetFullPath(Environment.CurrentDirectory);
            string extractedSection = ExtractSectionById(htmlFilePath, sectionId);
        if (extractedSection != null)
        {
            // Convert the extracted section to an image
            int imageWidth = 2240;
            int imageHeight = 2240;
            string base64Image = await HtmlToImage(extractedSection, imageWidth, imageHeight);

                // Load the HTML document
                HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.Load(htmlFilePath);

            // Find the section node by its ID
            var sectionNode = htmlDoc.GetElementbyId(sectionId);
            if (sectionNode != null)
            {
                // Create a new image node
                var imageNode = htmlDoc.CreateElement("img");
                imageNode.SetAttributeValue("src", "data:image/png;base64," + base64Image);
                imageNode.SetAttributeValue("alt", sectionId);

                // Replace the section node with the image node
                var parent = sectionNode.ParentNode;
                parent.ReplaceChild(imageNode, sectionNode);
                    var newhtml = parent.ParentNode.ParentNode.ParentNode;
                    File.WriteAllText(htmlFilePath, newhtml.InnerHtml);
                    // Save the modified HTML back to the file
                  //  htmlDoc.Save(htmlFilePath);

            }
        }
    }



}
}