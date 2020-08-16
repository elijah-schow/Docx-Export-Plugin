using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace DocxExportPlugin
{
    class Program
    {
        static void Main(string[] args)
        {
            string basePath = @"C:\Users\elija\Code\docx-export-plugin\";
            string inputPath = basePath + "Chinese Economy - Weak.fsx";
            string templatePath = basePath + "Template.docx";
            Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
            string outputPath = basePath +  "Output " + unixTimestamp + ".docx";

            var document = WordprocessingDocument.Open(templatePath, true);
            ImportFsx(inputPath, document);
            document.SaveAs(outputPath);

            Console.WriteLine("Saved " + outputPath);

        }

        public static void ImportFsx(string inputPath, WordprocessingDocument document)
        {
            Body body = document.MainDocumentPart.Document.Body;
            Paragraph p = null;
            Run r = null;
            string t = "";
            string style = "";
            bool ignore = false;

            // Start reading the fsx file
            FileStream file = File.Open(inputPath, FileMode.Open);
            BinaryReader reader = new BinaryReader(file, Encoding.ASCII);
            
            // Header
            string title = reader.ReadLine();

            string[] size = reader.ReadLine().Split(",");
            float width = float.Parse(size[0]);
            float height = float.Parse(size[1]);

            string[] margins = reader.ReadLine().Split(",");
            float left = float.Parse(margins[0]);
            float right = float.Parse(margins[1]);
            float top = float.Parse(margins[2]);
            float bottom = float.Parse(margins[3]);

            string styleSheet = reader.ReadLine();
            string characterStyleSheet = reader.ReadLine();
            string header = reader.ReadLine();
            string footer = reader.ReadLine();

            // Title
            if (title != "")
            {
                p = new Paragraph(new Run(new Text(title)));
                Utility.StyleParagraph(document, "Title", p);
                body.PrependChild(p);
                p = null;
            }

            // Body
            while(!reader.IsEndOfStream())
            {
                byte b = reader.ReadByte();

                // Start Codes
                if (b == 27 || b == 28) {
                    string code = new string(reader.ReadChars(2));
                    switch (code)
                    {
                        case "ST":
                            // Append the previous paragraph to the document
                            // TODO: DRY. This exact code is repeated in the newline section
                            if (p != null)
                            {
                                // End the previous run if it's still ongoing
                                if (r != null)
                                {
                                    // Ignore empty runs
                                    if (t != "")
                                    {
                                        r.AppendChild(new Text(t));
                                        t = "";

                                        p.AppendChild(r);
                                    }
                                    r = null;
                                }

                                // Ignore empty paragraphs
                                if (p.InnerText != "" && !ignore && style != "TOC Heading" && style != "Modify Date")
                                {
                                    body.AppendChild(p);
                                }
                                p = null;
                            }

                            // Create the next paragraph
                            p = new Paragraph();
                            style = ReadFSXString(reader);
                            Utility.StyleParagraph(document, style, p);

                            break;
                        case "SC":
                            if (r != null)
                            {
                                // Append text to the current run
                                r.AppendChild(new Text(t));
                                t = "";

                                // Append current run to the current paragraph
                                p.AppendChild(r);
                                r = null;
                            }

                            // Create the next run
                            r = new Run();
                            style = ReadFSXString(reader);
                            Utility.StyleRun(document, style, r);
                            break;
                        // Ignore Factsmith's table of contents
                        case "MI":
                            ignore = b == 27;
                            break;
                    }
                // ASCI Text
                } else if (b >= 32 && b <= 128)
                {
                    // Start a new run if one doesn't exist already
                    if (r == null)
                    {
                        r = new Run();
                    }
                    t += Convert.ToChar(b);
                // Newlines
                } else if (b == 10 || b == 13)
                {
                    // Skip the second line break if there are two in a row (e.g. \r\n)
                    int next = reader.PeekChar();
                    if (next == 10 || next == 13)
                    {
                        reader.ReadChar();
                    }

                    // Append the previous paragraph to the document
                    if (p != null)
                    {
                        // End the previous run if it's still ongoing
                        if (r != null)
                        {
                            // Ignore empty runs
                            if (t != "")
                            {
                                r.AppendChild(new Text(t));
                                t = "";

                                p.AppendChild(r);
                            }
                        }

                        // Ignore empty paragraphs and Factsmith's TOC
                        if (p.InnerText != "" && !ignore && style != "TOC Heading" && style != "Modify Date")
                        {
                            body.AppendChild(p);
                        }

                        // Create a new paragraph with the same style as before
                        style = p?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                        p = new Paragraph();
                        Utility.StyleParagraph(document, style, p);

                        // Create a new run with the same style as before
                        style = r?.RunProperties?.RunStyle?.Val;
                        r = new Run();
                        Utility.StyleRun(document, style, r);
                    }
                }
            }

            // TODO:
            // - [x] Character styles
            // - [x] Eliminate empty paragraphs
            // - [x] Mult-paragraph quotes
            // - [x] render Title
            // - [x] exclude factsmith modified date
            // - [x] exclude factsmith TOC
            // - [ ] render proper TOC
            // - [ ] Inline italics
            // - [ ] Notes
            // - [ ] Command line / Factsmith interface
            // - [ ] Images, bullets, etc?

            // - [ ] Factsmith template
        }

        public static string ReadFSXString(BinaryReader reader)
        {
            int length = reader.ReadByte() - 29;
            return new string(reader.ReadChars(length));
        }


    }
}
