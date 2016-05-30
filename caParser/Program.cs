using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using Fizzler.Systems.HtmlAgilityPack;
using Newtonsoft.Json.Linq;
using System.IO;
using Newtonsoft.Json;
using ClosedXML.Excel;
using System.Data;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Web;

namespace caParser
{
    class Program
    {
        static void Main(string[] args)
        {
            var headerBGColor = XLColor.FromArgb(174, 59, 36);
            var headerColor = XLColor.White;
            var oddColor = XLColor.FromArgb(224, 230, 196);
            var evenColor = XLColor.FromArgb(248, 247, 228);
           

            var currentDirectory = System.IO.Directory.GetCurrentDirectory();
            //ParseState("ca");
            var str = File.ReadAllText(Path.Combine(currentDirectory, "states.json"));
            //var states = JObject.Parse(str);
            var states = JsonConvert.DeserializeObject<List<USState>>(str);
            var outputDir = Path.Combine(currentDirectory, "Output");
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);
            //var xlOut = "REI.xlsx";
            var xlOut = "REI Club List.xlsx";
            var fileOut = Path.Combine(outputDir, xlOut);//"REI Club List.xlsx");
            ClosedXML.Excel.XLWorkbook book = new ClosedXML.Excel.XLWorkbook();
            var sheet = book.Worksheets.Add("REI Club List");
            sheet.Style.Font.FontName="Garamond";
            sheet.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sheet.Style.Font.FontSize = 11;
            sheet.Style.Font.Bold = true;
            

            sheet.Cell(1,1).Value = "REI CLUB LIST - STATEWISE";
            sheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            sheet.Cell("G1").Value = "STEWARD REDEVELOPMENT 2016";
            sheet.Cell("G1").Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Right;
            sheet.Cell("G1").Style.Font.Italic = true;

            sheet.Range("A1:G1").Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(83, 141, 213);
            sheet.Range("A1:G1").Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
            //sheet.Range("A1:G1").Style.Font.FontName = "Garamond";
            sheet.Range("A1:G1").Style.Font.FontSize = 26;
            sheet.Range("A1:G1").Style.Font.Bold = true;
            sheet.Row(2).Height = 4.5;
            sheet.Range("A2:G2").Style.Fill.BackgroundColor = headerBGColor;

            int startRow = 3;
            foreach (var usstate in states)
            //for (int i = 0; i < 4;i++)
            {
                //var usstate = states[i];
                Console.WriteLine("Getting data for {0}", usstate.name);
                var lst = ParseState(usstate.abbreviation.ToLower());
                if (lst != null)
                {
                    Console.WriteLine("Got {0} items", lst.Count);
                    //Console.WriteLine("Got {0} items", lst.Rows.Count);
                    //add two extra items
                    lst.Add(new Cre());
                    lst.Add(new Cre());

                    var tbl = sheet.Cell(startRow, 1).InsertTable<Cre>(lst);
                    tbl.ShowAutoFilter = false;

                    //email
                    foreach(var cell in tbl.Column(4).Cells())
                    {
                        var value = cell.Value.ToString();
                        if(!String.IsNullOrWhiteSpace(value))
                        {
                            Uri result;
                            if (Uri.TryCreate(value, UriKind.Absolute, out result))
                            {
                                cell.Value = value.Replace("mailto:", "").Trim();
                                cell.Hyperlink.ExternalAddress = result;
                            }

                        }
                    }
                    //website
                    foreach (var cell in tbl.Column(6).Cells())
                    {
                        var value = cell.Value.ToString();
                        if (!String.IsNullOrWhiteSpace(value))
                        {
                            Uri result;
                            if (Uri.TryCreate(value, UriKind.Absolute, out result))
                            {
                                //cell.Value = value.Replace("mailto:", "").Trim();
                                cell.Hyperlink.ExternalAddress = result;
                            }

                        }
                    }

                    tbl.Cell(1, 1).Value = usstate.name.ToUpper();
                    
                    
                    
                    //cells.

                    foreach (var row in tbl.Rows())
                    {
                        var rowNum = row.RowNumber();
                        row.Style.Fill.BackgroundColor = ((rowNum % 2) == 0) ? oddColor : evenColor;



                    }
                    var headerCells = tbl.Row(1).Cells();
                    headerCells.Style.Fill.BackgroundColor = headerBGColor;
                    headerCells.Style.Font.FontColor = XLColor.White;
                    headerCells.Style.Font.FontSize = 11.5;

                    startRow += lst.Count() + 1;

                }
                else
                {
                    Console.WriteLine("Got no items");
                }
            }
            sheet.Columns().AdjustToContents();
            book.SaveAs(fileOut);
            Console.WriteLine("Finished processing. Press any key to exit...");
            Console.ReadKey();
            System.Diagnostics.Process.Start(fileOut);
        }

        private static List<Cre> ParseState(string state)
        //private static DataTable ParseState(string state)
        {
            string url = String.Format("http://www.creonline.com/{0}.html", state);
            //string page = String.Empty;
            //using(WebClient client = new WebClient())
            //{
            //    page = client.DownloadString(url);
            //}
            //HtmlDocument doc = new HtmlDocument();
            //doc.LoadHtml(page);
            //var parentNode = doc.DocumentNode.SelectNodes("")
            List<HtmlNode> lstNodes = new List<HtmlNode>();
            List<Cre> lstCre = new List<Cre>();

            var web = new HtmlWeb();
            var document = web.Load(url);
            var page = document.DocumentNode;
            if (page.InnerHtml.Contains("Sorry but we can't find that page!"))
                return null;
            var tdNode = page.QuerySelector("td[valign='top']");

            if(tdNode!=null)
            {
                foreach (var pitem in tdNode.QuerySelectorAll("p"))
                {
                    lstNodes.AddRange(pitem.Descendants().Where(e=>e.Name.ToLower()!="br").Where(e=> !(e.Name.ToLower()=="#text" && String.IsNullOrWhiteSpace(e.InnerHtml))));
                }
            }
            var firstImg = lstNodes.First(e => e.Name.ToLower() == "img");
            for (int i = lstNodes.IndexOf(firstImg); i < lstNodes.Count; i++)
            {
                var node = lstNodes[i];
                if(node.Name.ToLower()=="img")
                {
                    lstCre.Add(new Cre());
                    
                }
                else
                {
                    var currentCre = lstCre.Last();
                    if(currentCre!=null)
                    {
                        string text = node.InnerHtml;

                        if (node.Name.ToLower() == "b" || node.Name.ToLower() == "strong")
                        {
                            if (node.FirstChild != null)
                            {
                                if (node.FirstChild.Name.ToLower() == "#text")
                                {
                                    currentCre.Name = node.InnerHtml;
                                }
                                if (node.FirstChild.Name.ToLower() == "a")
                                {
                                    currentCre.Name = node.FirstChild.InnerHtml;
                                }
                            }
                        }
                        if(node.Name.ToLower()=="a")
                        {
                            //check if url
                            var href = node.GetAttributeValue("href", "");
                            Uri result;
                            if(Uri.TryCreate(href,UriKind.Absolute,out result))
                            {
                                if (result.AbsoluteUri.ToLower().StartsWith("mailto"))
                                {
                                    currentCre.EmailAddress = result.AbsoluteUri;//.Replace("mailto:", "");
                                    currentCre.ContactPerson = node.InnerHtml;
                                    
                                }
                                else
                                {
                                    currentCre.Website = node.InnerHtml;
                                    
                                }
                            }
                        }
                        if(node.Name.ToLower()=="#text")
                        {
                            if (node.InnerHtml.Trim().StartsWith("Telephone:"))
                            {
                                currentCre.Phone = node.InnerHtml.Replace("Telephone:", "").Trim();
                                
                            }
                            if(node.InnerHtml.Trim().StartsWith("Where:"))
                            {
                                currentCre.Address = HttpUtility.HtmlDecode(node.InnerHtml.Replace("Where:", "").Trim());
                                
                            }
                            if(node.InnerHtml.Trim().StartsWith("Contact:"))
                            {
                                var contact = node.InnerHtml.Replace("Contact:", "");
                                if (!String.IsNullOrWhiteSpace(contact) && string.IsNullOrWhiteSpace(currentCre.ContactPerson))
                                {
                                    currentCre.ContactPerson = contact.Trim();
                                   
                                }
                                    
                            }
                        }
                    }
                }
            }
            var invalidCre = lstCre.FirstOrDefault(e => (!String.IsNullOrWhiteSpace(e.Name) && (e.Name.Trim().ToLower() == "click here")));
            if(invalidCre!=null)
            {
                lstCre.Remove(invalidCre);
            }

            return lstCre;
            
        }
    }
    class Cre
    {
        public string Name { get; set; }

        [Display(Name = "ADDRESS")]
        public string Address { get; set; }

        [Display(Name="PHONE #")]
        public string Phone { get; set; }

        [Display(Name = "EMAIL ADDRESS")]
        public string EmailAddress { get; set; }

        [Display(Name = "CONTACT PERSON")]
        public string ContactPerson { get; set; }

        [Display(Name = "WEBSITE")]
        public string Website { get; set; }

        [Display(Name = "NOTES")]
        public string Notes { get; set; }
        
        
        

    }
    class USState
    {
        public string name { get; set; }
        public string abbreviation { get; set; }
    }
}
