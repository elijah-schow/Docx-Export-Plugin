using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// TODO:
// - [x] Character styles
// - [x] Eliminate empty paragraphs
// - [x] Mult-paragraph quotes
// - [x] render Title
// - [x] exclude factsmith modified date
// - [x] exclude factsmith TOC
// - [x] Command line / Factsmith interface
// - [ ] Inline italics
// - [ ] Factsmith template
// - [ ] render proper TOC
// - [ ] Notes
// - [ ] Images, bullets, etc?
// - [ ] Handle file not found
namespace DocxExportPlugin
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get "Preserve input file" "/P" flag
            bool preserve = args.Contains("/P");

            // Get file paths
            string inputPath = args.Last();
            string outputPath = Path.ChangeExtension(inputPath, "docx");
            string templatePath = @"Template.docx";

            // Open the template file
            var document = WordprocessingDocument.Open(templatePath, true);

            // Start reading the fsx file
            FileStream file = File.Open(inputPath, FileMode.Open);
            BinaryReader reader = new BinaryReader(file);

            // Set up iteration variables
            Body body = document.MainDocumentPart.Document.Body;
            Paragraph p = null;
            Run r = null;
            string t = "";
            string style = "";
            bool ignore = false;

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
                                        r.AppendChild(new Text() { Text = t, Space = SpaceProcessingModeValues.Preserve });
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
                                // Only trim if this is an underlined character style
                                string cStyle = r?.RunProperties?.RunStyle?.Val;
                                if (t.Last() == ' ' && (cStyle == "Citation-ReadThis" || cStyle == "Quote-ReadThis"))
                                {
                                    // Append text to the current run and DO trim whitespace
                                    r.AppendChild(new Text(t));

                                    // Add one space to the beginning of the next run
                                    t = " ";
                                } else
                                {
                                    // Append text to the current run and DON'T trim whitespace
                                    r.AppendChild(new Text() { Text = t, Space = SpaceProcessingModeValues.Preserve });
                                    t = "";
                                }


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
                } else if (b >= 32)
                {
                    // Start a new run if one doesn't exist already
                    if (r == null)
                    {
                        r = new Run();
                    }
                    // Fixme: this is not reading extended ASCII correctly
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
                                r.AppendChild(new Text() { Text = t, Space = SpaceProcessingModeValues.Preserve });
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

            // Save and clean up
            file.Close();
            document.SaveAs(outputPath);

            // Delete the input file unless the "/P" flag is set
            if(!preserve)
            {
                File.Delete(inputPath);
            }

        }

        public static string ReadFSXString(BinaryReader reader)
        {
            int length = reader.ReadByte() - 29;
            return new string(reader.ReadChars(length));
        }


    }
}
